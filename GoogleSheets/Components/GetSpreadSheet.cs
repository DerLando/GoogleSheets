using System;
using System.Collections.Generic;
using Google.Apis.Sheets.v4;
using Grasshopper.Kernel;
using Newtonsoft.Json;
using Rhino.Geometry;

namespace GoogleSheets.Components
{
    public class GetSpreadSheet : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the GetSpreadSheet class.
        /// </summary>
        public GetSpreadSheet()
          : base("GetSpreadSheet", "GetSPRDSHT",
              "Gets a Spreadsheet for the specified Id",
              Settings.MainCategory, Settings.SubCategoryGet)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Id", "I", "Id of spreadsheet to get", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "S", "Spreadsheet gotten from server", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string id = "";
            if (!DA.GetData(0, ref id)) return;

            // get authorized service
            var service = Core.Authorizer.GetAuthorizedService();

            // build request object for id
            SpreadsheetsResource.GetRequest request = service.Spreadsheets.Get(id);
            request.IncludeGridData = true;

            // execute request, result is spreadsheet object
            var spreadSheet = request.Execute();

            // set data on output
            DA.SetData(0, spreadSheet);
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
            get { return new Guid("ea8c8978-a8bc-49db-999d-548b8eb9ede6"); }
        }
    }
}