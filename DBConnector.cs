using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace MaterialTransfer
{
  public class DBConnector
  {
    string[] scopes = { SheetsService.Scope.Spreadsheets };
    public string applicationname = "Material transfer";
    string spreadsheetId = "1LTUA8C-tIBgieNw7PjDGjkSBvV0_i5vpwnJc2C_sKLk";
    private SheetsService service;
    private GoogleCredential credential;

    public DBConnector()
    {
      string connecitonSecret = Assembly.GetExecutingAssembly().GetName().Name + ".MaterialSecret.json";
      using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(connecitonSecret))
      {
        credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
      }

      service = new SheetsService(new BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = applicationname
      });
    }

    public IList<IList<object>> ReadSheetData(string range)
    {
      var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
      ValueRange response = request.Execute();

      return response.Values;
    }
  }
}
