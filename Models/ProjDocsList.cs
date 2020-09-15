using Microsoft.SharePoint.Client;
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
            ViewTitle = "All Documents";
            ColumnForDefaultView = new List<string>
            {
                "DocDescription",
                "DocType"
            };
        }
    }
}
