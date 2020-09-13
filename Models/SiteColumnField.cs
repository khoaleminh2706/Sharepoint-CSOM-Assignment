using Microsoft.SharePoint.Client;
using System;

namespace CreateSPSite.Models
{
    public class SiteColumnField: AbstractField
    {
        public SiteColumnField(
            ClientContext clientContext)
            :base(clientContext)
        {
        }

        public override void Create()
        {
            if (TargetContentType == null) throw new Exception("Cần cung cấp content type");
            
            Field targetField = GetField();

            if (targetField == null)
            {
                throw new Exception($"Field {InternalName} không tồn tại.");
            }

            FieldLinkCreationInformation fldLink = new FieldLinkCreationInformation
            {
                Field = targetField
            };

            fldLink.Field.Required = false;
            fldLink.Field.Hidden = false;

            TargetContentType.FieldLinks.Add(fldLink);
            TargetContentType.Update(false);
            _context.ExecuteQuery();
        }
    }
}
