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
        }

        protected override List AddCustomColum(List list)
        {
            return null;
        }

        protected override List AddView()
        {
            return null;
        }
    }
}
