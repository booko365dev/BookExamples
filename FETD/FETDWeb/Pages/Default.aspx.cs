using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FETDWeb
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        //gavdcodebegin 02
        protected void Page_Load(object sender, EventArgs e)
        {
            var spContext = 
                        SharePointContextProvider.Current.GetSharePointContext(Context);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                clientContext.Load(clientContext.Web, web => web.Title);
                clientContext.ExecuteQuery();
                Response.Write(clientContext.Web.Title);
            }

            string[] allQstring = Request.QueryString.AllKeys;
            string myString = string.Empty;
            foreach (string oneQstring in allQstring)
            {
                string oneValue = Request.QueryString[oneQstring];
                string oneKey = oneQstring.Trim();
                myString += oneKey + " - " + oneValue + "<br />";
            }
            Response.Write(myString);
        }
        //gavdcodeend 02
    }
}