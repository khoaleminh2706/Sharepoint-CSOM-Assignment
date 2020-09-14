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

        public void CreateEmployeeContentType()
        {
                ContentTypeCollection contentTypeCollection;
                contentTypeCollection = _clientContext.Web.ContentTypes;

                _clientContext.Load(contentTypeCollection);
                _clientContext.ExecuteQuery();

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

                    _clientContext.Load(item);
                    _clientContext.ExecuteQuery();

                    Console.WriteLine("Add column....");

                    Field targetField = _clientContext.Web.AvailableFields.GetByInternalNameOrTitle("FirstName");

                    FieldLinkCreationInformation fldLink = new FieldLinkCreationInformation();
                    fldLink.Field = targetField;

                    // If uou set this to "true", the column getting added to the content type will be added as "required" field
                    fldLink.Field.Required = false;

                    // If you set this to "true", the column getting added to the content type will be added as "hidden" field
                    fldLink.Field.Hidden = false;

                    item.FieldLinks.Add(fldLink);
                    item.Update(false);
                    _clientContext.ExecuteQuery();

                    Console.WriteLine("Add column finished....");

                    Console.WriteLine("Finish creating Content Type");
                }

                Console.WriteLine("Creating list...");

                // Access subsite
                Web hRWeb = _clientContext.Site.OpenWeb("HR");

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

                _clientContext.Load(newList);
                _clientContext.ExecuteQuery();

                contentTypeCollection = newList.ContentTypes;

                _clientContext.Load(contentTypeCollection);
                _clientContext.ExecuteQuery();

                _clientContext.Load(newList);
                _clientContext.ExecuteQuery();

                // Update the view
                View view = newList.Views.GetByTitle("All Items");
                _clientContext.Load(view, v => v.ViewFields);
                Field name = newList.Fields.GetByInternalNameOrTitle("FirstName");

                _clientContext.Load(name);
                _clientContext.ExecuteQuery();

                view.ViewFields.Add(name.InternalName);
                view.Update();
                _clientContext.ExecuteQuery();

                // Execute the query to the server.
                _clientContext.ExecuteQuery();

                Console.WriteLine("Finished creating list...");
        }

        /// <summary>
        /// Tạo Project list
        /// </summary>
        public void CreateProjectList()
        { 
                Web rootWeb = _clientContext.Site.RootWeb;
                ContentTypeCollection contentTypeCollection = _clientContext.Web.ContentTypes;

                _clientContext.Load(contentTypeCollection);
                _clientContext.ExecuteQuery();

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

                    _clientContext.Load(item);
                    _clientContext.ExecuteQuery();

                    Console.WriteLine("Add column....");

                    string projectNameFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Text' Name='Project Name' StaticName='ProjectName' DisplayName='Project Name' />";
                    Field projectNameField = rootWeb.Fields.AddFieldAsXml(projectNameFieldSchema, false, AddFieldOptions.AddFieldInternalNameHint);
                    projectNameField.Group = "Training";
                    item.FieldLinks.Add(new FieldLinkCreationInformation
                    {
                        Field = projectNameField,
                    });

                    item.Update(false);
                    _clientContext.ExecuteQuery();

                    Console.WriteLine("Add column finished....");

                    Console.WriteLine("Finish creating Content Type");
                }

                Console.WriteLine("Creating list...");

                // Access subsite
                Web hRWeb = _clientContext.Site.OpenWeb("HR");

                // Find Employees list
                _clientContext.Load(hRWeb.Lists);
                _clientContext.ExecuteQuery();

                var employeesList = hRWeb.Lists.GetByTitle("Employees");
                _clientContext.Load(employeesList);
                _clientContext.ExecuteQuery();

                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = "Projects";
                creationInfo.Description = "New list description";
                creationInfo.TemplateType = (int)ListTemplateType.GenericList;

                List newList = hRWeb.Lists.Add(creationInfo);
                newList.ContentTypesEnabled = true;
                newList.ContentTypes.AddExistingContentType(item);
                
                _clientContext.Load(newList);

                contentTypeCollection = newList.ContentTypes;

                _clientContext.Load(contentTypeCollection);

                // Remove Item
                ContentType targetContentType = (from contentType in contentTypeCollection where contentType.Name == "Item" select contentType).FirstOrDefault();

                if (targetContentType != null)
                {
                    targetContentType.DeleteObject();
                }

                string leaderFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Lookup' Name='Leader' StaticName='Leader' DisplayName='Leader' List='" + employeesList.Id + "' ShowField='Title' />";
                Field leaderField = newList.Fields.AddFieldAsXml(leaderFieldSchema, false, AddFieldOptions.AddToDefaultContentType);
                _clientContext.Load(leaderField);
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
                _clientContext.Load(view, v => v.ViewFields);
                Field name = newList.Fields.GetByInternalNameOrTitle("ProjectName");
                Field leader = newList.Fields.GetByInternalNameOrTitle("Leader");

                _clientContext.Load(name);
                _clientContext.Load(leader);
                _clientContext.ExecuteQuery();

                view.ViewFields.Add(name.InternalName);
                view.Update();
                _clientContext.ExecuteQuery();

                // Execute the query to the server.
                _clientContext.ExecuteQuery();

                Console.WriteLine("Finished creating list...");
        }
    }
}
