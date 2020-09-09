using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Security;

namespace CreateSPSite
{
    public static class SharepointService
    {
        private const string loginName = "khoa.le.minh@khoaleminh.onmicrosoft.com";
        private const string password = "P3t3rLeMinh";
        const string ITFirm = "https://khoaleminh.sharepoint.com/sites/ITFirmpub";

        public static void CreateEmployeeContentType()
        {
            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(ITFirm))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);


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

                    Console.WriteLine("Creating list...");

                    // Access subsite
                    Web hRWeb = clientContext.Site.OpenWeb("HR");

                    ListCreationInformation creationInfo = new ListCreationInformation();
                    creationInfo.Title = "Employees";
                    creationInfo.Description = "New list description";
                    creationInfo.TemplateType = (int)ListTemplateType.GenericList;

                    List newList = hRWeb.Lists.Add(creationInfo);
                    newList.ContentTypes.AddExistingContentType(newContentType);

                    contentTypeCollection = newList.ContentTypes;

                    clientContext.Load(contentTypeCollection);
                    clientContext.ExecuteQuery();

                    ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "Item" select contentType).FirstOrDefault();

                    if (targetContentType != null)
                    {
                        targetContentType.DeleteObject();
                    }

                    clientContext.Load(newList);
                    clientContext.ExecuteQuery();

                    // Update the view
                    View view = newList.Views.GetByTitle("AllItems");
                    clientContext.Load(view, v => v.ViewFields);
                    Field name = newList.Fields.GetByInternalNameOrTitle("FirstName");
                    
                    clientContext.Load(name);
                    clientContext.ExecuteQuery();

                    view.ViewFields.Add(name.InternalName);
                    view.Update();
                    clientContext.ExecuteQuery();

                    // Execute the query to the server.
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Finished creating list...");
                }
                else
                {
                    Console.WriteLine("Content type đã tồn tạo....");
                }
            }
        }
    }
}
