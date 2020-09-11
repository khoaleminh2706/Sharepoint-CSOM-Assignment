using Microsoft.Online.SharePoint.TenantAdministration;
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
        const string ITFirm = "https://khoaleminh.sharepoint.com/sites/newsite1";

        public static void CreateEmployeeContentType(string siteUrl)
        {
            Console.WriteLine("Create Employees list");
            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));

            using (ClientContext clientContext = new ClientContext(siteUrl))
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
                newList.ContentTypes.AddExistingContentType(item);

                clientContext.Load(newList);
                clientContext.ExecuteQuery();

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
                View view = newList.Views.GetByTitle("All Items");
                clientContext.Load(view, v => v.ViewFields);
                Field firstName = newList.Fields.GetByInternalNameOrTitle("FirstName");

                clientContext.Load(firstName);
                clientContext.ExecuteQuery();

                view.ViewFields.Add(firstName.InternalName);
                view.Update();
                clientContext.ExecuteQuery();

                // Execute the query to the server.
                clientContext.ExecuteQuery();

                Console.WriteLine("Finished creating list...");
            }
        }

        public static string CreateSite(string adminSiteUrl, string rootSiteUrl, string siteTitle, string siteUrl)
        {
            var secureString = new SecureString();
            password.ToCharArray().ToList().ForEach(c => secureString.AppendChar(c));
            
            siteUrl = rootSiteUrl + "/sites/" + siteUrl;

            using (ClientContext clientContext = new ClientContext(adminSiteUrl))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);

                #region Create Site
                try
                {
                    var tenant = new Tenant(clientContext);
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

                    clientContext.Load(spo);
                    Console.WriteLine("Start creating site...");
                    clientContext.ExecuteQuery();

                    //Check if provisioning of the SiteCollection is complete.
                    while (!spo.IsComplete)
                    {
                        //Wait for 30 seconds and then try again
                        System.Threading.Thread.Sleep(30000);
                        //spo.RefreshLoad();
                        clientContext.Load(spo);
                        Console.WriteLine("Sau 30 giây....");
                        clientContext.ExecuteQuery();
                    }

                    Console.WriteLine("Site Created.");
                    
                    Console.WriteLine("Creating HR subsite");

                    // Get new create web
                    var subclientContext = new ClientContext(siteUrl);
                    subclientContext.Credentials = new SharePointOnlineCredentials(loginName, secureString);

                    WebCreationInformation webCreationInfo = new WebCreationInformation
                    {
                        Url = "HR",
                        Title = "HR Department",
                        Description = "Subsite for HR",
                        UseSamePermissionsAsParentSite = true,
                        WebTemplate  = "STS#0",
                        Language = 1033,
                    };
                    subclientContext.Site.RootWeb.Webs.Add(webCreationInfo);
                    subclientContext.ExecuteQuery();

                    // add HR site link to quick lauch menu
                    var quickLaunchNav = subclientContext.Web.Navigation.QuickLaunch;
                    
                    NavigationNodeCreationInformation newNode = new NavigationNodeCreationInformation
                    {
                        Title = "HR",
                        Url = siteUrl + "/HR",
                        AsLastNode = true
                    };
                    quickLaunchNav.Add(newNode);
                    subclientContext.Load(quickLaunchNav);
                    subclientContext.ExecuteQuery();

                    Console.WriteLine("Finising creating HR subsite...");
                    return siteUrl;
                }
                catch (ServerException ex)
                {
                    Console.WriteLine($"Lỗi: {ex.Message}");
                    return null;
                }
                #endregion
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

        public static void CreateDocumentList()
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

                ContentType newContentType = (from contentType in contentTypeCollection where contentType.Name == "Project Document" select contentType).FirstOrDefault();

                if (newContentType != null)
                {
                    Console.WriteLine("Content type already exists....");
                }
                else
                {
                    Console.WriteLine("Start creating content type.");

                    ContentType parentContentType = (from contentType in contentTypeCollection where contentType.Name == "Document" select contentType).FirstOrDefault();

                    ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation
                    {
                        Name = "Project Document",
                        // Description of the new content type
                        Description = "Project Document",

                        // Name of the group under which the new content type will be creted
                        Group = "Training",
                        ParentContentType = parentContentType
                    };

                    newContentType = contentTypeCollection.Add(contentTypeCreationInformation);

                    clientContext.Load(newContentType);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Add column....");

                    Field targetField = clientContext.Web.AvailableFields.GetByInternalNameOrTitle("Description");

                    FieldLinkCreationInformation fldLink = new FieldLinkCreationInformation();
                    fldLink.Field = targetField;

                    // If uou set this to "true", the column getting added to the content type will be added as "required" field
                    fldLink.Field.Required = false;

                    // If you set this to "true", the column getting added to the content type will be added as "hidden" field
                    fldLink.Field.Hidden = false;

                    newContentType.FieldLinks.Add(fldLink);
                    newContentType.Update(false);
                    clientContext.ExecuteQuery();

                    string projectNameFieldSchema = @"<Field ID='" + Guid.NewGuid() + "' Type='Choice' Name='DocType' StaticName='DocType' DisplayName='Document Type' Format='Dropdown' Group='Training' FillInChoice='FALSE' >" 
                        + "<Default>Business requirement</Default>"
                        +"<CHOICES>" 
                        + "<CHOICE>Business requirement</CHOICE>" 
                        + "<CHOICE>Technical document</CHOICE>" 
                        + "<CHOICE>User guide</CHOICE></CHOICES></Field>";
                    Field projectNameField = rootWeb.Fields.AddFieldAsXml(projectNameFieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                    newContentType.FieldLinks.Add(new FieldLinkCreationInformation
                    {
                        Field = projectNameField,
                    });

                    newContentType.Update(false);
                    clientContext.ExecuteQuery();

                    Console.WriteLine("Add column finished....");

                    Console.WriteLine("Finish creating Content Type");
                }

                Console.WriteLine("Creating list...");

                // Access subsite
                Web hRWeb = clientContext.Site.OpenWeb("HR");

                var projectList = hRWeb.Lists.GetByTitle("Projects");
                clientContext.Load(projectList);
                clientContext.ExecuteQuery();

                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = "Project Documents";
                creationInfo.Description = "Projects Doc Library";
                creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;

                List projDocList = hRWeb.Lists.Add(creationInfo);
                projDocList.ContentTypesEnabled = true;
                projDocList.ContentTypes.AddExistingContentType(newContentType);

                clientContext.Load(projDocList);

                contentTypeCollection = projDocList.ContentTypes;

                clientContext.Load(contentTypeCollection);
                clientContext.ExecuteQuery();

                // Remove Item
                //ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "Document" select contentType).FirstOrDefault();

                //if (targetContentType != null)
                //{
                //    targetContentType.DeleteObject();
                //}

                string projFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Lookup' Name='Proj' StaticName='Proj' DisplayName='Project' List='" + projectList.Id + "' ShowField='Title' />";
                Field projField = projDocList.Fields.AddFieldAsXml(projFieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                projField.SetShowInEditForm(true);
                projField.SetShowInNewForm(true);
                clientContext.Load(projField);

                projDocList.Update();
                clientContext.ExecuteQuery();

                // Update the view
                View view = projDocList.Views.GetByTitle("All Documents");
                clientContext.Load(view, v => v.ViewFields);
                Field desc = projDocList.Fields.GetByInternalNameOrTitle("Description");
                Field doctype = projDocList.Fields.GetByInternalNameOrTitle("DocType");
                
                clientContext.Load(desc);
                clientContext.Load(doctype);
                clientContext.ExecuteQuery();

                view.ViewFields.Add(desc.InternalName);
                view.ViewFields.Add(doctype.InternalName);
                view.ViewFields.Add(projField.InternalName);
                view.Update();
                clientContext.ExecuteQuery();

                // Execute the query to the server.
                clientContext.ExecuteQuery();

                Console.WriteLine("Finished creating list...");
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
    }
}
