using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Security;

namespace CreateSPSite
{
    class Program
    {
        static void Main()
        {
            string loginName = "khoa.le.minh@khoaleminh.onmicrosoft.com";
            string password = "P3t3rLeMinh";

            const string HRSiteUrl = "https://khoaleminh.sharepoint.com/sites/NormalITfirm/HR/";
            const string ITFirm = "https://khoaleminh.sharepoint.com/sites/ITFirmpub";

            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(HRSiteUrl))
            {
                #region Authen
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);
                #endregion

                #region Add column to list
                List targetList = clientContext.Web.Lists.GetByTitle("Employees");

                string schemaRichTextField = "<Field ID='" + Guid.NewGuid() + "' Type='Note' Name='Comments' StaticName='Comments' DisplayName = 'Comments' NumLines = '6' RichText = 'TRUE' RichTextMode = 'FullHtml' IsolateStyles = 'TRUE' Sortable = 'FALSE' /> ";
                Field fileToAdd = targetList.Fields.AddFieldAsXml(schemaRichTextField, true, AddFieldOptions.AddFieldInternalNameHint);

                clientContext.Load(fileToAdd);
                clientContext.ExecuteQuery();
                #endregion
            }
        }
    }
}
