using Microsoft.SharePoint.Client;

namespace CreateSPSite.Models
{
    public class EmployeeContentType : AbstractContentType
    {
        public EmployeeContentType(ClientContext clientContext): base(clientContext)
        {
            Name = "Employee";
            Parent = "Contact";
        }

        public override void Create()
        {
            base.Create();
        }
    }
}
