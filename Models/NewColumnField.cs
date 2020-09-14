using Microsoft.SharePoint.Client;
using System;

namespace CreateSPSite.Models
{
    public class NewColumnField : AbstractField
    {
        public NewColumnField(
            ClientContext context): base(context)
        {
        }

        public override void Create()
        {
            if (TargetContentType == null) throw new Exception("Cần cung cấp content type");

            var rootWeb = _context.Site.RootWeb;

            Field newField = rootWeb.Fields.AddFieldAsXml(XmlSchema, false, AddFieldOptions.AddFieldInternalNameHint);
            newField.Group = "Training";

            TargetContentType.FieldLinks.Add(new FieldLinkCreationInformation
            {
                Field = newField,
            });

            TargetContentType.Update(false);
            _context.ExecuteQuery();
        }
    }
}
