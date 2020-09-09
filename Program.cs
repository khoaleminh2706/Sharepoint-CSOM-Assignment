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

            const string siteUrl = "https://khoaleminh-admin.sharepoint.com/";
            const string adminSiteUrl = "https://khoaleminh-admin.sharepoint.com/";
            const string rootSiteUrl = "https://khoaleminh.sharepoint.com/sites/";
            const string HRSiteUrl = "https://khoaleminh.sharepoint.com/sites/NormalITfirm/HR/";
            const string ITFirm = "https://khoaleminh.sharepoint.com/sites/ITFirmpub";

            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(ITFirm))
            {
                #region Authen
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);
                #endregion

                ContentTypeCollection contentTypeCollection;
                contentTypeCollection = clientContext.Web.ContentTypes;

                clientContext.Load(contentTypeCollection);
                clientContext.ExecuteQuery();

                ContentType item = (from contentType in contentTypeCollection where contentType.Name == "Employee" select contentType).FirstOrDefault();
                clientContext.Load(contentTypeCollection, ctColl => ctColl.Include(ct => ct.Name).Where(ct => ct.Name == "Employee"));
                clientContext.ExecuteQuery();

                if (contentTypeCollection.Count == 0) 
                {
                    Console.WriteLine("Bắt đầu tạo content type.");
                    ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation
                    {
                        Name = "Employee",
                        // Description of the new content type
                        Description = "New Content Type Description",

                        // Name of the group under which the new content type will be creted
                        Group = "Training"
                    };

                    // Add "ContentTypeCreationInformation" object created above
                    ContentType newContentType = contentTypeCollection.Add(contentTypeCreationInformation);

                    clientContext.Load(newContentType);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Tạo xong");
                    
                    Console.WriteLine("Add column....");

                    Field targetField = clientContext.Web.AvailableFields.GetByInternalNameOrTitle("FirstName");

                    FieldLinkCreationInformation fldLink = new FieldLinkCreationInformation();
                    fldLink.Field = targetField;

                    // If uou set this to "true", the column getting added to the content type will be added as "required" field
                    fldLink.Field.Required = false;

                    // If you set this to "true", the column getting added to the content type will be added as "hidden" field
                    fldLink.Field.Hidden = false;

                    newContentType.FieldLinks.Add(fldLink);
                    newContentType.Update(false);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Add column finished....");
                }
                else
                {
                    Console.WriteLine("Content type đã tồn tạo....");
                }
            }
        }
    }
}
