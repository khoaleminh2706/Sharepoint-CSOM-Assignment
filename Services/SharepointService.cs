using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;

namespace CreateSPSite.Services
{
    public class SharepointService
    {
        private readonly ClientContext _context;

        public SharepointService(ClientContext clientContext)
        {
            _context = clientContext;
        }

        public Web CheckHRSubsiteExist()
        {
            Web hrWeb = _context.Site.OpenWeb("HR");
            _context.Load(hrWeb);
            _context.ExecuteQuery();
            return hrWeb;
        }

        public string CreateSite(string rootSiteUrl, string loginName, string siteTitle, string siteUrl)
        {
            siteUrl = rootSiteUrl + "/sites/" + siteUrl;

            #region Create Site
            var tenant = new Tenant(_context);
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

            _context.Load(spo);
            Console.WriteLine("Start creating site...");
            _context.ExecuteQuery();

            //Check if provisioning of the SiteCollection is complete.
            while (!spo.IsComplete)
            {
                //Wait for 30 seconds and then try again
                System.Threading.Thread.Sleep(30000);
                //spo.RefreshLoad();
                _context.Load(spo);
                Console.WriteLine("Sau 30 giây....");
                _context.ExecuteQuery();
            }

            Console.WriteLine("Site Created.");
            return siteUrl;
            #endregion
        }

        public Web CreateHRSubsite()
        {
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

                Web web = _context.Site.RootWeb.Webs.Add(webCreationInfo);
                _context.Load(web);
                _context.ExecuteQuery();
                return web;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi tạo HR subsite");
                Console.WriteLine("Lỗi: " + ex.GetType().Name + " " + ex.Message);
            }
            return null;
        }
    }
}
