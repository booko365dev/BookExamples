using System;

namespace RBYP
{
    public partial class GenerateGuid : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        //gavdcodebegin 002
        //ATTENTION: Replaced by TZGQ 
        protected void btnGenerateGuid_Click(object sender, EventArgs e)
        {
            lblNewGuid.Text = Guid.NewGuid().ToString();
        }
        //gavdcodeend 002
    }
}