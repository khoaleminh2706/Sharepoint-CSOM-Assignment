using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CreateSPSite.Models
{
    public abstract class AbstractContentType: IDisposable
    {
        private ClientContext _context;
        public string Name { get; set; }
        public string Description { get; set; } = "New Custom Content Type";
        public string Group { get; set; } = "Training";
        public List<AbstractField> FieldsList { get; set; }

        public AbstractContentType(ClientContext context)
        {
            _context = context;
        }

        public virtual ContentType Create()
        {
            ContentTypeCollection contentTypeColl = _context.Web.ContentTypes;
            _context.Load(contentTypeColl);
            _context.ExecuteQuery();

            var targetContentType = CheckExits(contentTypeColl);

            if (targetContentType != null)
            {
                throw new Exception("Content type already Exists");
            }
            else
            {
                // Create content Type
                ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation
                {
                    Name = Name,
                    Description = Description,
                    Group = Group
                };

                targetContentType = contentTypeColl.Add(contentTypeCreationInformation);

                _context.Load(targetContentType);
                _context.ExecuteQuery();
                return targetContentType;
            }
        }

        private ContentType CheckExits(ContentTypeCollection collection)
        {
            return (
                from contentType in collection 
                where contentType.Name == Name 
                select contentType)
                .FirstOrDefault();
        }

        public void CreateFields(ContentType targetContextType)
        {
            if (FieldsList != null && FieldsList.Count != 0)
                FieldsList.ForEach(field =>
                {
                    field.TargetContentType = targetContextType;
                    field.Create();
                });
        }

        public void Dispose()
        {
            _context.Dispose();
        }
    }
}
