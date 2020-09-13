using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace CreateSPSite.Models
{
    public abstract class AbstractContentType: IDisposable
    {
        private ClientContext _context;
        public string Name { get; set; }
        public string Description { get; set; } = "New Custom Content Type";
        public string Group { get; set; } = "Training";

        public AbstractContentType(ClientContext context)
        {
            _context = context;
        }

        public virtual void Create()
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
                    // Description of the new content type
                    Description = Description,

                    // Name of the group under which the new content type will be creted
                    Group = Group
                };

                targetContentType = contentTypeColl.Add(contentTypeCreationInformation);

                _context.Load(targetContentType);
                _context.ExecuteQuery();
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

        public void Dispose()
        {
            _context.Dispose();
        }
    }
}
