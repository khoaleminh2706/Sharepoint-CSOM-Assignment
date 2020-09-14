using Microsoft.SharePoint.Client;
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
            ViewTitle = "All Items";
            ColumnList = new List<string>
            {
                "ProjectName",
                "ProjDescription",
                "State",
                "StartDate",
                "_EndDate"
            };
        }
    }
}
