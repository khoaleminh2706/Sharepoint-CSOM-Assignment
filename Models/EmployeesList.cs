using Microsoft.SharePoint.Client;

namespace CreateSPSite.Models
{
    public class EmployeesList : AbstractList
    {
        public EmployeesList(ClientContext context): base(context)
        {
            Title = Constants.ListTitle.Employees;
            ContentTypeName = Constants.ContentType.Employee;
            TemplateType = (int)ListTemplateType.GenericList;
        }
    }
}
