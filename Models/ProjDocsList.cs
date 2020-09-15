using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace CreateSPSite.Models
{
    public class ProjDocsList : AbstractList
    {
        public ProjDocsList(ClientContext context) : base(context)
        {
            Title = Constants.ListTitle.ProjDoc;
            ContentTypeTitle = Constants.ContentType.ProjectDoc;
            TemplateType = (int)ListTemplateType.DocumentLibrary;
            DependListTitle = Constants.ListTitle.Projects;
            ViewTitle = "All Documents";
            ColumnForDefaultView = new List<string>
            {
                "DocDescription",
                "DocType"
            };
        }

        protected override List AddCustomColum(List list)
        {
            var lookupList = CheckListExists(_context.Web.Lists, DependListTitle);
            if (lookupList == null)
                throw new Exception($"List {DependListTitle} không tồn tại");
            _context.Load(lookupList, li => li.Id);
            _context.ExecuteQuery();
            
            string projectFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Lookup' Name='" + DependListTitle + "' StaticName='Project' DisplayName='Project' List='" + lookupList.Id + "' ShowField='ProjectName' />";
            Field projField = list.Fields.AddFieldAsXml(projectFieldSchema, true, AddFieldOptions.AddFieldInternalNameHint);
            projField.SetShowInEditForm(true);
            projField.SetShowInNewForm(true);
            _context.Load(projField);

            list.Update();
            return list;
        }
    }
}
