using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CreateSPSite.Models
{
    public abstract class AbstractContentType: IDisposable
    {
        private ClientContext _context;

        public AbstractContentType(ClientContext context)
        {
            _context = context;
        }

        #region Methods
        public virtual ContentType Create()
        {
            ContentTypeCollection contentTypeColl = _context.Web.ContentTypes;
            _context.Load(contentTypeColl);
            _context.ExecuteQuery();

            var targetContentType = GetContentType(contentTypeColl);

            if (targetContentType != null)
            {
                throw new Exception("Content type already Exists");
            }
            else
            {
                var parentContentType = GetContentType(contentTypeColl, ParentTypeTitle);
                // Create content Type
                ContentTypeCreationInformation contentTypeCreationInformation = new ContentTypeCreationInformation
                {
                    Name = Name,
                    Description = Description,
                    Group = Group,
                    ParentContentType = parentContentType
                };

                targetContentType = contentTypeColl.Add(contentTypeCreationInformation);

                _context.Load(targetContentType);
                _context.ExecuteQuery();
                return targetContentType;
            }
        }

        private ContentType GetContentType(ContentTypeCollection collection, string contentTypeTitle = "")
        {
            contentTypeTitle = contentTypeTitle != "" ? contentTypeTitle : Name;
            return (
                from contentType in collection 
                where contentType.Name == contentTypeTitle 
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
        #endregion

        #region Properties
        public string Name { get; set; }
        public string Description { get; set; } = "New Custom Content Type";
        public string Group { get; set; } = "Training";
        public List<AbstractField> FieldsList { get; set; }
        public string ParentTypeTitle { get; set; } = "Item";
        #endregion
    }
}
