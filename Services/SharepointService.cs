using Microsoft.Online.SharePoint.SPLogger;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace CreateSPSite.Services
{
    public class SharepointService
    {
        private readonly ClientContext _clientContext;

        public SharepointService(ClientContext clientContext)
        {
            _clientContext = clientContext;
        }

        public Web CheckHRSubsiteExist()
        {
            Web hrWeb = _clientContext.Site.OpenWeb("HR");
            _clientContext.Load(hrWeb);
            _clientContext.ExecuteQuery();
            return hrWeb;
        }

        public string CreateSite(string adminSiteUrl, string rootSiteUrl, string loginName, string siteTitle, string siteUrl)
        {
            siteUrl = rootSiteUrl + "/sites/" + siteUrl;

            #region Create Site
            try
            {
                var tenant = new Tenant(_clientContext);
                var siteCreationProperties = new SiteCreationProperties
                {

                    //New SiteCollection Url
                    Url = siteUrl,

                    //Title of the Root Site
                    Title = siteTitle,

                    //Login name of Owner
                    Owner = loginName,

                    //Template of the Root Site. Using Team Site for now.
                    // BLANKINTERNETCONTAINER#0 STS#0
                    Template = "BLANKINTERNETCONTAINER#0",

                    //Storage Limit in MB
                    StorageMaximumLevel = 5,

                    TimeZoneId = 7
                };

                //Create the SiteCollection
                SpoOperation spo = tenant.CreateSite(siteCreationProperties);

                _clientContext.Load(spo);
                Console.WriteLine("Start creating site...");
                _clientContext.ExecuteQuery();

                //Check if provisioning of the SiteCollection is complete.
                while (!spo.IsComplete)
                {
                    //Wait for 30 seconds and then try again
                    System.Threading.Thread.Sleep(30000);
                    //spo.RefreshLoad();
                    _clientContext.Load(spo);
                    Console.WriteLine("Sau 30 giây....");
                    _clientContext.ExecuteQuery();
                }

                Console.WriteLine("Site Created.");
                return siteUrl;
            }
            catch (ServerException ex)
            {
                Console.WriteLine($"Lỗi: {ex.Message}");
                return null;
            }
            #endregion
        }

        public string CreateHRSubsite()
        {
            string resultUrl = "";
            Console.WriteLine("Creating HR subsite");
            try
            {
                WebCreationInformation webCreationInfo = new WebCreationInformation
                {
                    Url = "HR",
                    Title = "HR Department",
                    Description = "Subsite for HR",
                    UseSamePermissionsAsParentSite = true,
                    WebTemplate = "STS#0",
                    Language = 1033,
                };

                Web web = _clientContext.Site.RootWeb.Webs.Add(webCreationInfo);
                _clientContext.Load(web);
                _clientContext.ExecuteQuery();
                resultUrl = web.Url;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi tạo HR subsite");
                Console.WriteLine("Lỗi: " + ex.GetType().Name + " " + ex.Message);
            }
            return resultUrl;
        }
    }
}
