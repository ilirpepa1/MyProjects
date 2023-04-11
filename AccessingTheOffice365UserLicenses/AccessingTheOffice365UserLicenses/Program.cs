
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using Newtonsoft.Json;
using System.Reflection;

namespace AccessingTheOffice365UserLicenses
{
    class Program
    {
        static async Task Main(string[] args)
        {       
            string executedAt = getStructuredTime(DateTime.Now);
            string executionfolderPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            executionfolderPath = executionfolderPath.Substring(6);
            string fullExcelFilePath = executionfolderPath;

            string logFilePath = executionfolderPath + @"\log-" + executedAt;// + ".txt";
            System.IO.File.AppendAllText(logFilePath, "Iniziato l'esecuzione di console" + Environment.NewLine);
            string errorLogFilePath = executionfolderPath + @"\ErrorLog-" + executedAt;// + ".txt";

            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create("a96e6d89-c66e-4055-84c8-e6cf8b6523f1")//ClientId
                          .WithRedirectUri("https://myapp.com")
                          .WithClientSecret("p~07Q~-HtTJIqnpnvJ.rlB11vCq3H0LPpnQgI") //ClientSecret
                          .WithAuthority(new Uri("https://login.microsoftonline.com/e4874dc3-6e49-4e01-bc68-81f70c984cff/"))//Instance + Tenant
                          .Build();

            string[] scopes = new string[] { $"https://graph.microsoft.com/.default" }; //ApiUrl + .default
            GraphServiceClient graphServiceClient =
            new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
            {
                // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                var authResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                // Add the access token in the Authorization header of the API
                requestMessage.Headers.Authorization =
                new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
            })
            );
            List<User> users = new List<User>();
            // Get the first page
            IGraphServiceUsersCollectionPage usersPage = await graphServiceClient.Users.Request().GetAsync();

            // Add the first page of results to the user list
            users.AddRange(usersPage.CurrentPage);

            // Fetch each page and add those results to the list
            while (usersPage.NextPageRequest != null)
            {
                usersPage = await usersPage.NextPageRequest.GetAsync();
                users.AddRange(usersPage.CurrentPage);
            }

            Microsoft.Identity.Client.AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            if (result != null)
            {
                Dictionary<string, List<string>> UsersAndTheirOwnedLicenses = new Dictionary<string, List<string>>();
                foreach (var user in users)
                {
                    string UserPrincipalName = user.UserPrincipalName;
                    string UserDisplayName = user.DisplayName;
                   
                    await GetLicenseAssigned($"https://graph.microsoft.com/v1.0/users", result.AccessToken, UserPrincipalName, UsersAndTheirOwnedLicenses, UserDisplayName, errorLogFilePath);      
                }
                WriteToExcel(UsersAndTheirOwnedLicenses);
            }
        }
        public static async Task GetLicenseAssigned(string webApiUrl, string accessToken, string UserPrincipalName, Dictionary<string, List<string>> UsersAndTheirOwnedLicenses, string UserDisplayName, string errorLogFilePath)
        {
            var httpClient = new HttpClient();
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequestHeaders = httpClient.DefaultRequestHeaders;
                if (defaultRequestHeaders.Accept == null || !defaultRequestHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                string webApiUrlOneUSer1 = webApiUrl + "/" + UserPrincipalName + "/licenseDetails";
                HttpResponseMessage response = await httpClient.GetAsync(webApiUrlOneUSer1);

                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject(json) as JObject;
                    List<string> allLicencesOfUser = StructuredLincenses(result);
                    try
                    {
                        if (allLicencesOfUser.Count != 0)
                        {
                            UsersAndTheirOwnedLicenses.Add(UserPrincipalName, allLicencesOfUser);
                        }
                        else
                        {
                            UsersAndTheirOwnedLicenses.Add(UserPrincipalName, null);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.IO.File.AppendAllText(errorLogFilePath, $"Utente {UserPrincipalName} e' stato trovato piu di una volta. {ex.Message}" + Environment.NewLine);
                        Console.WriteLine($"Utente {UserPrincipalName} e' stato trovato piu di una volta. {ex.Message}");
                    }                
                }
            }
        }
        private static List<string> StructuredLincenses(JObject result)
        {
            List<string> allLicencesOfUser = new List<string>();
            foreach (JProperty child in result.Properties().Where(p => !p.Name.StartsWith("@")))
            {
                foreach (var item in child.Value)
                {
                    foreach (var i in item)
                    {
                        var actualLicense = ((Newtonsoft.Json.Linq.JProperty)i).Name == "skuPartNumber" ? ((Newtonsoft.Json.Linq.JProperty)i).First : null;
                        if (actualLicense != null)
                        {
                            allLicencesOfUser.Add(i.First.ToString());          
                        }
                    }
                }
            }
            return allLicencesOfUser;
        }        
 
        public static void WriteToExcel(Dictionary<string, List<string>> UsersAndTheirOwnedLicenses) 
        {
            Microsoft.Office.Interop.Excel.Application myexcelApplication = new Microsoft.Office.Interop.Excel.Application();
            if (myexcelApplication != null)
            {
                Microsoft.Office.Interop.Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet myexcelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)myexcelWorkbook.Sheets.Add();
                myexcelWorksheet.Cells[1, 1] = "User";
                myexcelWorksheet.Cells[1, 2] = "User's License";
                
                int i = 2;
                foreach (var item in UsersAndTheirOwnedLicenses)
                {
                    myexcelWorksheet.Cells[i, 1] = item.Key;
                    if (item.Value != null)
                    {
                        for (int j = 0; j < item.Value.Count; j++)
                        {
                            myexcelWorksheet.Cells[i, 2] = item.Value[j];
                            i++;
                        }
                    }
                    else
                    {
                        myexcelWorksheet.Cells[i, 2] = "No License";
                        i++;
                    }
                    i++;
                }

                myexcelWorkbook.Worksheets[myexcelWorkbook.Sheets.Count].Delete();
                Console.WriteLine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + '\\' + "UsersAndTheirOwnedLicenses.xls");
                myexcelApplication.ActiveWorkbook.SaveAs(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location)  + '\\' + "UsersAndTheirOwnedLicenses.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                myexcelWorkbook.Close();
                myexcelApplication.Quit();
            }
        }

        public static string getStructuredTime(DateTime dt)
        {
            return dt.Year + "." + dt.Month + "." + dt.Day + "_" + dt.Hour + "hh " + dt.Minute + "mm " + dt.Second + "ss";
        }
        public static async Task MethodGetAllLicences(string webApiUrl, string accessToken, List<Licenses> allLicences)
        {
            var httpClient = new HttpClient();
            if (!string.IsNullOrEmpty(accessToken))
            {
                var defaultRequestHeaders = httpClient.DefaultRequestHeaders;
                if (defaultRequestHeaders.Accept == null || !defaultRequestHeaders.Accept.Any(m => m.MediaType == "application/json"))
                {
                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                }
                defaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // GetAsync method get
                HttpResponseMessage response = await httpClient.GetAsync(webApiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject(json) as JObject;
                    Test_StructuredLincenses(result, allLicences);
                }
            }
        }
        private static void Test_StructuredLincenses(JObject result, List<Licenses> allLicences)
        {
            foreach (JProperty child in result.Properties().Where(p => !p.Name.StartsWith("@")))
            {
                foreach (var item in child.Value)
                {
                    Licenses ul = new Licenses();
                    ul.NameofLicense = item.First.First.ToString();
                    ul.IdOfLicense = item.Last.Last.ToString();
                    allLicences.Add(ul);
                }
            }
        }
    }

    public class Licenses
    {
        public string NameofLicense;
        public string IdOfLicense;
    }
}
   
