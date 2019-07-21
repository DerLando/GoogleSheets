using System;
using System.Collections.Generic;
using Google.Apis.Sheets.v4.Data;
using Grasshopper.Kernel;
using Rhino.Geometry;

namespace GoogleSheets.Components
{
    public class ExplodeSpreadSheet : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ExplodeSpreadSheet class.
        /// </summary>
        public ExplodeSpreadSheet()
          : base("ExplodeSpreadSheet", "SPRDSHTBOOM",
              "Explodes a spreadsheet into its parts",
              Settings.MainCategory, Settings.SubCategoryExplode)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "S", "Spreadsheet to explode", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Title", "T", "Title of spreadsheet", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheets", "S", "Sheets in spreadsheet", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Spreadsheet spreadSheet = null;
            if (!DA.GetData(0, ref spreadSheet)) return;

            DA.SetData(0, spreadSheet.Properties.Title);
            DA.SetDataList(1, spreadSheet.Sheets);
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
            get { return new Guid("f8c02c0f-d783-4093-bb3c-8ac6fe456ecd"); }
        }
    }
}