using Microsoft.SharePoint.Client;
using System.Collections.Generic;

namespace CreateSPSite.Models
{
    public class EmployeeContentType : AbstractContentType
    {
        public EmployeeContentType(ClientContext clientContext): base(clientContext)
        {
            Name = "Employee";
            FieldsList = new List<AbstractField>()
            {
                new SiteColumnField(clientContext) { InternalName = "FirstName" }
            };
        }
    }
}
