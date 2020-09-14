using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace CreateSPSite.Models
{
    public abstract class AbstractList: IDisposable
    {
        private ClientContext _context;

        public AbstractList(ClientContext context)
        {
            _context = context;
        }

        #region Methods
        public virtual List Create()
        {
            // check list already exists
            _context.Load(_context.Web.Lists);
            _context.Load(_context.Site.RootWeb.ContentTypes);
            _context.ExecuteQuery();

            var list = CheckExists(_context.Web.Lists);
           
            if (list != null)
                throw new Exception("List đã tồn tại");

            // Check content type exists
            var targetContentType = GetContentType(_context.Site.RootWeb.ContentTypes);
            if (targetContentType == null)
                throw new Exception($"Content Type {ContentTypeName} không tồn tại. Vui lòng tạo content type trước.");

            ListCreationInformation creationInfo = new ListCreationInformation
            {
                Title = Title,
                Description = "New list description",
                TemplateType = TemplateType
            };

            List newList = _context.Web.Lists.Add(creationInfo);
            newList.ContentTypesEnabled = true;
            newList.ContentTypes.AddExistingContentType(targetContentType);

            _context.Load(newList);
            _context.ExecuteQuery();
            return newList;
        }

        protected List CheckExists(ListCollection collection, string listTitle = "")
        {
            listTitle = listTitle != "" ? listTitle : Title;
            return (from list in collection where list.Title == listTitle select list)
                .FirstOrDefault();
        }

        protected ContentType GetContentType(ContentTypeCollection collection)
            => (from contentType in collection where contentType.Name == ContentTypeName select contentType)
                .FirstOrDefault();

        public void Dispose()
        {
            _context.Dispose();
        }
        #endregion

        #region Properties
        public string Title { get; set; }
        public string ContentTypeName { get; set; }
        public int TemplateType { get; set; } = (int)ListTemplateType.GenericList;
        #endregion
    }
}
