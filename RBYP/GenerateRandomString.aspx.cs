using System;

namespace RBYP
{
    public partial class GenerateRandomString : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btnGenerateRandomString_Click(object sender, EventArgs e)
        {
            lblRandomString.Text =
                        System.IO.Path.GetRandomFileName().Replace(".", string.Empty);
        }
    }
}
