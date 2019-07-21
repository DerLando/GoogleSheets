using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace GoogleSheets
{
    public class GoogleSheetsInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "GoogleSheets";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return null;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("4133c9b3-b0f5-4de9-8ba3-b880a7571628");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "";
            }
        }
    }
}
