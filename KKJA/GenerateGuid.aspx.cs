using System;

namespace KKJA
{
    public partial class GenerateGuid : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btnGenerateGuid_Click(object sender, EventArgs e)
        {
            lblNewGuid.Text = Guid.NewGuid().ToString();
        }
    }
}