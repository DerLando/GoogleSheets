using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Grasshopper.Kernel;
using Rhino.Geometry;

namespace GoogleSheets
{
    public class GetSheetCell : GH_Component
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";


        /// <summary>
        /// Initializes a new instance of the GetSheetCell class.
        /// </summary>
        public GetSheetCell()
          : base("GetSheetCell", "Nickname",
              "Description",
              "GoogleSheets", "Subcategory")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet Id", "Id", "The spreadsheet id", GH_ParamAccess.item);
            pManager.AddTextParameter("Sheet Name", "N", "The sheet name", GH_ParamAccess.item);
            pManager.AddTextParameter("Cell Name", "C", " ", GH_ParamAccess.item);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Cell Text", "T", "Cell text", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string spreadsheet_Id = "";
            string sheetName = "";
            string cellName = "";


            if (!DA.GetData(0, ref spreadsheet_Id)) return;
            if (!DA.GetData(1, ref sheetName)) return;
            if (!DA.GetData(2, ref cellName)) return;

            string cellText = GetCellText(spreadsheet_Id, sheetName, cellName);

            DA.SetData(0, cellText);


        }

        private string GetCellText(string spreadsheetId, string sheetName, string cellName)
        {
            // If modifying these scopes, delete your previously saved credentials
            // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String range = String.Format("{0}!{1}", sheetName, cellName);
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(spreadsheetId, range);
            

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                return values[0][0].ToString();
            }
            return "Error";
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("cacb10f5-05d8-46e5-ae27-4e4c92b8d814"); }
        }
    }
}