using Microsoft.SharePoint.Client;

namespace CreateSPSite.Models
{
    public class EmployeesList : AbstractList
    {
        public EmployeesList(ClientContext context): base(context)
        {
            Title = Constants.ListTitle.Employees;
            ContentTypeTitle = Constants.ContentType.Employee;
            TemplateType = (int)ListTemplateType.GenericList;
            ViewTitle = "All Items";
            ColumnList = new System.Collections.Generic.List<string>
            {
                "FirstName",
                "ShortDesc",
                "ProgrammingLanguages"
            };
        }
    }
}
