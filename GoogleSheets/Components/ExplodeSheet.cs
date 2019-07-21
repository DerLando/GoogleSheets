using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4.Data;
using GoogleSheets.Core;
using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Rhino.Geometry;

namespace GoogleSheets.Components
{
    public class ExplodeSheet : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ExplodeSheet class.
        /// </summary>
        public ExplodeSheet()
          : base("ExplodeSheet", "SHTBOOM",
              "Explodes a sheet in its components",
              Settings.MainCategory, Settings.SubCategoryExplode)
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet to explode", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Title", "T", "Title of sheet", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Id", "Id", "Id of sheet", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Index", "I", "Index of sheet", GH_ParamAccess.item);
            pManager.AddTextParameter("Rows", "R", "Height of sheet", GH_ParamAccess.tree);

        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            Sheet sheet = null;
            if (!DA.GetData(0, ref sheet)) return;

            DA.SetData(0, sheet.Properties.Title);
            DA.SetData(1, sheet.Properties.SheetId);
            DA.SetData(2, sheet.Properties.Index);

            // empty tree for cell strings
            DataTree<string> rowsTree = new DataTree<string>();
            

            // extract strings, add to tree
            var rowData = sheet.Data[0].RowData;
            for (int i = 0; i < rowData.Count; i++)
            {
                // create new branch
                rowsTree.Insert("", new GH_Path(i), 0);
            }
            // clear dummy data
            rowsTree.ClearData();

            for (int i = 0; i < rowData.Count; i++)
            {
                List<string> rowStrings = new List<string>();
                if (!(rowData[i].Values is null))
                {
                    foreach (var value in rowData[i].Values)
                    {
                        //rowStrings.Add(value.EffectiveValue?.ToDataString());
                        rowStrings.Add(value.FormattedValue);
                    }
                }
                rowsTree.Branch(i).AddRange(rowStrings);
            }

            // set row data
            DA.SetDataTree(3, rowsTree);
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
            get { return new Guid("147856a5-ed00-4e0b-be12-376ed23ea78f"); }
        }
    }
}