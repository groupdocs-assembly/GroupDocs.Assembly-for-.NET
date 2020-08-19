using System;
using System.Web.UI;

namespace GroupDocs.Assembly.Live.Demos.UI.Assembly
{
    public partial class Default : System.Web.UI.Page
    {
        protected FileFormat fileFormat;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Page.RouteData.Values["extension"] != null)
            {
                string extension = Page.RouteData.Values["extension"].ToString().ToLower();
            }
        }
    }
}