using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Grasshopper.Kernel;
using Rhino.Geometry;

namespace GoogleSheets
{
    public class SpreadsheetIdFromUrl : GH_Component
    {
        public static string IdRegexPattern = "/spreadsheets/d/([a-zA-Z0-9-_]+)";

        /// <summary>
        /// Initializes a new instance of the SpreadsheetIdFromUrl class.
        /// </summary>
        public SpreadsheetIdFromUrl()
          : base("SpreadsheetIdFromUrl", "URL -> Id",
              "Description",
              "GoogleSheets", "Subcategory")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet Url", "URL", " ", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Spreadsheet Id", "Id", " ", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string url = " ";

            if (!DA.GetData(0, ref url)) return;

            string id = GetId(url);
            if (string.IsNullOrEmpty(id)) return;

            DA.SetData(0, id);
        }

        private string GetId(string url)
        {
            string regResult = Regex.Match(url, IdRegexPattern, RegexOptions.CultureInvariant, new TimeSpan(250)).Value;
            if (regResult.Contains("/spreadsheets/d/"))
            {
                return regResult.Substring(16);
            }
            else return "";
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
            get { return new Guid("01da23dc-95e9-4f46-8ee0-48029e7fdb95"); }
        }
    }
}