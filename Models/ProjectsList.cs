using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace CreateSPSite.Models
{
    public class ProjectsList : AbstractList
    {
        public ProjectsList(ClientContext context): base(context)
        {
            Title = Constants.ListTitle.Projects;
            ContentTypeTitle = Constants.ContentType.Project;
            TemplateType = (int)ListTemplateType.GenericList;
            DependListTitle = Constants.ListTitle.Employees;
            ViewTitle = "All Items";
            ColumnForDefaultView = new List<string>
            {
                "ProjectName",
                "ProjDescription",
                "State",
                "StartDate",
                "_EndDate"
            };
        }

        protected override List AddCustomColum(List list)
        {
            var employeesList = CheckListExists(_context.Web.Lists, DependListTitle);
            if (employeesList == null)
                throw new Exception("List employee không tồn tại");
            _context.Load(employeesList, li => li.Id);
            _context.ExecuteQuery();
            
            string leaderFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='Lookup' Name='Leader' StaticName='Leader' DisplayName='Leader' List='" + employeesList.Id + "' ShowField='Title' />";
            Field leaderField = list.Fields.AddFieldAsXml(leaderFieldSchema, true, AddFieldOptions.AddFieldInternalNameHint);
            leaderField.SetShowInEditForm(true);
            leaderField.SetShowInNewForm(true);
            _context.Load(leaderField);

            // Add member field
            string memberFieldSchema = "<Field ID='" + Guid.NewGuid() + "' Type='LookupMulti' Name='Member' StaticName='Member' DisplayName='Member' List='" + employeesList.Id + "' ShowField='Title' Mult='TRUE' />";
            Field memberField = list.Fields.AddFieldAsXml(memberFieldSchema, true, AddFieldOptions.AddFieldInternalNameHint);
            memberField.SetShowInEditForm(true);
            memberField.SetShowInNewForm(true);
            _context.Load(memberField);

            list.Update();
            return list;
        }
    }
}
