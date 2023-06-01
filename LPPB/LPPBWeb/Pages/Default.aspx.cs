using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace LPPBWeb
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

        //gavdcodebegin 002
        protected void Page_Load(object sender, EventArgs e)
        {
            var spContext = SharePointContextProvider.Current.
                                                           GetSharePointContext(Context);

            using (ClientContext spClientContext =
                                            spContext.CreateUserClientContextForSPHost())
            {
                spClientContext.Load(spClientContext.Web, web => web.Title);
                spClientContext.ExecuteQuery();
                Response.Write(spClientContext.Web.Title);

                GetSiteInformation(spClientContext);
            }
        }
        //gavdcodeend 002

        //gavdcodebegin 003
        private void GetSiteInformation(ClientContext spClientContext)
        {
            if (IsPostBack)
            {
                _ = new Uri(Request.QueryString["SPHostUrl"]);
            }

            Web myWeb = spClientContext.Web;
            spClientContext.Load(myWeb);
            spClientContext.ExecuteQuery();

            spClientContext.Load(myWeb.CurrentUser);
            spClientContext.ExecuteQuery();
            string myCurrentUser = spClientContext.Web.CurrentUser.LoginName;

            ListCollection allLists = myWeb.Lists;
            spClientContext.Load<ListCollection>(allLists);
            spClientContext.ExecuteQuery();

            UserCollection allUsers = myWeb.SiteUsers;
            spClientContext.Load<UserCollection>(allUsers);
            spClientContext.ExecuteQuery();

            List<string> myUsers = new List<string>();
            foreach (User oneUser in allUsers)
            {
                myUsers.Add(oneUser.LoginName);
            }

            List<string> myLists = new List<string>();
            foreach (List oneList in allLists)
            {
                myLists.Add(oneList.Title);
            }

            lblUser.Text = myCurrentUser;
            lstOtherUsers.DataSource = myUsers;
            lstOtherUsers.DataBind();
            lstLists.DataSource = myLists;
            lstLists.DataBind();
        }
        //gavdcodeend 003
    }
}