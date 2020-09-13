using Microsoft.SharePoint.Client;

namespace CreateSPSite.Models
{
    public class ProjectContentType : AbstractContentType
    {
        public ProjectContentType(ClientContext clientContext) : base(clientContext)
        {
            Name = "Project";
        }

        public override void Create()
        {
            base.Create();
        }
    }
}
