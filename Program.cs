using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Security;

namespace CreateSPSite
{
    class Program
    {
        static void Main(string[] args)
        {
            string loginName = "khoa.le.minh@khoaleminh.onmicrosoft.com";
            string password = "P3t3rLeMinh";
            string siteUrl = "https://khoaleminh-admin.sharepoint.com/";
            string adminSiteUrl = "https://khoaleminh-admin.sharepoint.com/";
            string rootSiteUrl = "https://khoaleminh.sharepoint.com/sites/";

            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext("https://khoaleminh.sharepoint.com/sites/NormalITfirm/HR/"))
            {
                #region Authen
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);

                var site = clientContext.Site;
                var web = clientContext.Web;
                clientContext.Load(web);
                clientContext.Load(site);
                clientContext.ExecuteQuery();

                string rootSiteCollectionURL = site.ServerRelativeUrl;
                string SiteCollectionURL = web.Url;

                Console.WriteLine("Root site: " + rootSiteCollectionURL);
                Console.WriteLine("Site collection: " + SiteCollectionURL);

                #endregion
            }
            Console.ReadKey();
        }
    }
}
