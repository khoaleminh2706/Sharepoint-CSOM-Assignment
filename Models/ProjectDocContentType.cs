using Microsoft.SharePoint.Client;

namespace CreateSPSite.Models
{
    public class ProjectDocContentType : AbstractContentType
    {
        public ProjectDocContentType(ClientContext clientContext) : base(clientContext)
        {
            Name = "Employee";
        }
    }
}
