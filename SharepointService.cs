using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Security;

namespace CreateSPSite
{
    public static class SharepointService
    {
        private const string loginName = "khoa.le.minh@khoaleminh.onmicrosoft.com";
        private const string password = "P3t3rLeMinh";
        const string ITFirm = "https://khoaleminh.sharepoint.com/sites/newsite1";

        public static void CreateEmployeeContentType()
        {
            Console.WriteLine("Create Employees list");
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

                if (item != null)
                {
                    Console.WriteLine("Content type already exists....");
                }
                else
                {
                    Console.WriteLine("Start creating content type.");

                    ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation
                    {
                        Name = "Employee",
                        // Description of the new content type
                        Description = "New Content Type Description",

                        // Name of the group under which the new content type will be creted
                        Group = "Training"
                    };


                    item = contentTypeCollection.Add(contentTypeCreationInformation);

                    clientContext.Load(item);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Add column....");

                    Field targetField = clientContext.Web.AvailableFields.GetByInternalNameOrTitle("FirstName");

                    FieldLinkCreationInformation fldLink = new FieldLinkCreationInformation();
                    fldLink.Field = targetField;

                    // If uou set this to "true", the column getting added to the content type will be added as "required" field
                    fldLink.Field.Required = false;

                    // If you set this to "true", the column getting added to the content type will be added as "hidden" field
                    fldLink.Field.Hidden = false;

                    item.FieldLinks.Add(fldLink);
                    item.Update(false);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Add column finished....");

                    Console.WriteLine("Finish creating Content Type");
                }

                Console.WriteLine("Creating list...");

                // Access subsite
                Web hRWeb = clientContext.Site.OpenWeb("HR");

                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = "Employees";
                creationInfo.Description = "New list description";
                creationInfo.TemplateType = (int)ListTemplateType.GenericList;

                List newList = hRWeb.Lists.Add(creationInfo);
                newList.ContentTypesEnabled = true;

                // Delete Item Content Type
                ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "Item" select contentType).FirstOrDefault();
                if (targetContentType != null)
                {
                    targetContentType.DeleteObject();
                }

                // Add content type
                newList.ContentTypes.AddExistingContentType(item);

                clientContext.Load(newList);
                clientContext.ExecuteQuery();

                contentTypeCollection = newList.ContentTypes;

                clientContext.Load(contentTypeCollection);
                clientContext.ExecuteQuery();

                clientContext.Load(newList);
                clientContext.ExecuteQuery();

                // Update the view
                View view = newList.Views.GetByTitle("All Items");
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
        }

        /// <summary>
        /// Tạo Project list
        /// </summary>
        public static void CreateProjectList()
        {
            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(ITFirm))
            { 
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);
                
                Web rootWeb = clientContext.Site.RootWeb;
                ContentTypeCollection contentTypeCollection = clientContext.Web.ContentTypes;

                clientContext.Load(contentTypeCollection);
                clientContext.ExecuteQuery();

                ContentType item = (from contentType in contentTypeCollection where contentType.Name == "Project" select contentType).FirstOrDefault();

                if (item != null)
                {
                    Console.WriteLine("Content type already exists....");
                }
                else
                {
                    Console.WriteLine("Start creating content type.");

                    ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation
                    {
                        Name = "Project",
                        // Description of the new content type
                        Description = "New Content Type Description",

                        // Name of the group under which the new content type will be creted
                        Group = "Training"
                    };

                    item = contentTypeCollection.Add(contentTypeCreationInformation);

                    clientContext.Load(item);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Add column....");

                    string projectNameFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Text' Name='Project Name' StaticName='ProjectName' DisplayName='Project Name' />";
                    Field projectNameField = rootWeb.Fields.AddFieldAsXml(projectNameFieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                    projectNameField.Group = "Training";
                    item.FieldLinks.Add(new FieldLinkCreationInformation
                    {
                        Field = projectNameField,
                    });

                    item.Update(false);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Add column finished....");

                    Console.WriteLine("Finish creating Content Type");
                }

                Console.WriteLine("Creating list...");

                // Access subsite
                Web hRWeb = clientContext.Site.OpenWeb("HR");

                // Find Employees list
                clientContext.Load(hRWeb.Lists);
                clientContext.ExecuteQuery();

                var employeesList = hRWeb.Lists.GetByTitle("Employees");
                clientContext.Load(employeesList);
                clientContext.ExecuteQuery();

                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = "Projects";
                creationInfo.Description = "New list description";
                creationInfo.TemplateType = (int)ListTemplateType.GenericList;

                List newList = hRWeb.Lists.Add(creationInfo);
                newList.ContentTypesEnabled = true;
                newList.ContentTypes.AddExistingContentType(item);
                
                clientContext.Load(newList);

                contentTypeCollection = newList.ContentTypes;

                clientContext.Load(contentTypeCollection);

                // Remove Item
                ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "Item" select contentType).FirstOrDefault();

                if (targetContentType != null)
                {
                    targetContentType.DeleteObject();
                }

                string leaderFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Lookup' Name='Leader' StaticName='Leader' DisplayName='Leader' List='" + employeesList.Id + "' ShowField='Title' />";
                Field leaderField = newList.Fields.AddFieldAsXml(leaderFieldSchema, false, AddFieldOptions.AddToDefaultContentType);
                clientContext.Load(leaderField);
                leaderField.SetShowInEditForm(true);
                leaderField.SetShowInNewForm(true);
                leaderField.Update();

                // Add member field
                //string memberFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Lookup' Name='Member' StaticName='Member' DisplayName='Member' List='" + employeesList.Id + "' ShowField='Title' />";
                //Field memberField = rootWeb.Fields.AddFieldAsXml(leaderFieldSchema, true, AddFieldOptions.AddFieldInternalNameHint);

                //memberField = newList.Fields.Add(memberField);
                //memberField.SetShowInEditForm(true);
                //memberField.SetShowInNewForm(true);

                newList.Update();

                // Update the view
                View view = newList.Views.GetByTitle("All Items");
                clientContext.Load(view, v => v.ViewFields);
                Field name = newList.Fields.GetByInternalNameOrTitle("ProjectName");
                Field leader = newList.Fields.GetByInternalNameOrTitle("Leader");

                clientContext.Load(name);
                clientContext.Load(leader);
                clientContext.ExecuteQuery();

                view.ViewFields.Add(name.InternalName);
                view.Update();
                clientContext.ExecuteQuery();

                // Execute the query to the server.
                clientContext.ExecuteQuery();

                Console.WriteLine("Finished creating list...");
            }
        }

        /// <summary>
        /// Delete Content Type by Title
        /// </summary>
        /// <param name="name"></param>
        public static void DeleteContentType(string name)
        {
            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(ITFirm))
            {
                 clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);
                ContentTypeCollection oContentTypeCollection = clientContext.Web.ContentTypes;

                // Load content type collection
                clientContext.Load(oContentTypeCollection);
                clientContext.ExecuteQuery();

                ContentType targetContentType = (from contentType in oContentTypeCollection where contentType.Name == name select contentType).FirstOrDefault();

                // Delete Content Type
                targetContentType.DeleteObject();

                clientContext.ExecuteQuery();
            }
        }

        public static void FindContentTypeAssoc(string name)
        {
            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(ITFirm))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);
                ContentTypeCollection contentTypeColl = clientContext.Web.ContentTypes;

                clientContext.Load(contentTypeColl);
                clientContext.Load(clientContext.Web);
                clientContext.Load(clientContext.Web.Lists);
                clientContext.Load(clientContext.Web.Webs);
                clientContext.ExecuteQuery();

                foreach (var list in clientContext.Web.Lists)
                {
                    clientContext.Load(list.ContentTypes);
                    clientContext.ExecuteQuery();

                    var targetContentType = (from contentType in contentTypeColl where contentType.Name == name select contentType).FirstOrDefault();
                    if (targetContentType != null)
                    {
                        Console.WriteLine("Found at " + list.Title);
                    }
                }

                if (clientContext.Web.Webs.Count > 0)
                {
                    foreach (var web in clientContext.Web.Webs)
                    {
                        contentTypeColl = web.ContentTypes;
                        clientContext.Load(contentTypeColl);
                        clientContext.Load(web.Lists);
                        clientContext.ExecuteQuery();

                        foreach (var list in web.Lists)
                        {
                            clientContext.Load(list.ContentTypes);
                            clientContext.ExecuteQuery();

                            var targetContentType = (from contentType in contentTypeColl where contentType.Name == name select contentType).FirstOrDefault();
                            if (targetContentType != null)
                            {
                                Console.WriteLine("Found at " + list.Title);
                            }
                        }
                    }
                }
            }
        }


        public static void CreateProjectList1()
        {
            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(ITFirm))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);
                ContentTypeCollection contentTypeCollection = clientContext.Web.ContentTypes;

                clientContext.Load(contentTypeCollection);
                clientContext.ExecuteQuery();
                Console.WriteLine("Creating list...");

                ContentType item = (from contentType in contentTypeCollection where contentType.Name == "Project" select contentType).FirstOrDefault();

                // Access subsite
                Web hRWeb = clientContext.Site.OpenWeb("HR");

                // Find Employees list
                clientContext.Load(hRWeb.Lists);
                clientContext.ExecuteQuery();

                var employeesList = hRWeb.Lists.GetByTitle("Employees");
                clientContext.Load(employeesList);
                clientContext.ExecuteQuery();

                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = "Projects";
                creationInfo.Description = "New list description";
                creationInfo.TemplateType = (int)ListTemplateType.GenericList;

                List newList = hRWeb.Lists.Add(creationInfo);
                newList.ContentTypesEnabled = true;
                newList.ContentTypes.AddExistingContentType(item);

                clientContext.Load(newList);

                contentTypeCollection = newList.ContentTypes;

                clientContext.Load(contentTypeCollection);
                clientContext.ExecuteQuery();

                // Remove Item
                ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "Item" select contentType).FirstOrDefault();

                if (targetContentType != null)
                {
                    targetContentType.DeleteObject();
                }

                string leaderFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Lookup' Name='Leader' StaticName='Leader' DisplayName='Leader' List='" + employeesList.Id + "' ShowField='Title' />";
                Field leaderField = newList.Fields.AddFieldAsXml(leaderFieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                leaderField.SetShowInEditForm(true);
                leaderField.SetShowInNewForm(true);
                clientContext.Load(leaderField);

                // Add member field
                string memberFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='LookupMulti' Name='Member' StaticName='Member' DisplayName='Member' List='" + employeesList.Id + "' ShowField='Title' Mult='TRUE' />";
                Field memberField = newList.Fields.AddFieldAsXml(memberFieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                memberField.SetShowInEditForm(true);
                memberField.SetShowInNewForm(true);
                clientContext.Load(memberField);

                newList.Update();
                clientContext.ExecuteQuery();

                // Update the view
                View view = newList.Views.GetByTitle("All Items");
                clientContext.Load(view, v => v.ViewFields);
                Field name = newList.Fields.GetByInternalNameOrTitle("ProjectName");
                
                clientContext.Load(name);
                clientContext.ExecuteQuery();

                view.ViewFields.Add(name.InternalName);
                view.ViewFields.Add(leaderField.InternalName);
                view.ViewFields.Add(memberField.InternalName);
                view.Update();
                clientContext.ExecuteQuery();

                // Execute the query to the server.
                clientContext.ExecuteQuery();

                Console.WriteLine("Finished creating list...");
            }
        }
        // TODO: Delete List
    }
}
