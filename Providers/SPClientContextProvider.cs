using Microsoft.SharePoint.Client;
using System.Linq;
using System.Security;

namespace CreateSPSite.Provider
{
    public sealed class SPClientContextProvider
    {
        private ClientContext _context;
        
        public SPClientContextProvider(
            string loginName, 
            string password,
            string siteUrl = ""
            )
        {
            SiteUrl = siteUrl;
            LoginName = loginName;
            Password = password;
        }

        public ClientContext Create()
        {
            _context = new ClientContext(SiteUrl);
            var secureString = new SecureString();
            Password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));
            _context.Credentials = new SharePointOnlineCredentials(LoginName, secureString);

            return _context;
        }

        #region Properties
        public string SiteUrl { get; set; }
        public string LoginName { get; set; }
        public string Password { get; set; }
        #endregion
    }
}
