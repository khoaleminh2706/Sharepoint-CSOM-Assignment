using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace CreateSPSite.Models
{
    public abstract class AbstractList: IDisposable
    {
        private ClientContext _context;

        public AbstractList(ClientContext context, string name)
        {
            _context = context;
            Name = name;
        }

        #region Methods
        public virtual List Create()
        {
            //var employeesList = _context.Web.Lists.GetByTitle("Employees");
            //_context.Load(employeesList);
            //_context.ExecuteQuery();
            //return employeesList;

            // check list already exists
            _context.Load(_context.Web.Lists);
            _context.Load(_context.Web.ContentTypes);

            var list = CheckExists(_context.Web.Lists);
            if (list != null)
            {
                throw new Exception("List đã tồn tại");
            }

            // Check content type exists
            var targetContentType = CheckConentTypeExits(_context.Web.ContentTypes);
            if (targetContentType == null)
                throw new Exception("Content Type không tồn tại. Vui lòng tạo content type trước.");

            ListCreationInformation creationInfo = new ListCreationInformation
            {
                Title = Name,
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

        public List CheckExists(ListCollection collection)
        {
            return collection.GetByTitle(Name);
        }

        public ContentType CheckConentTypeExits(ContentTypeCollection collection)
        {
            return (from contentType in collection where contentType.Name == ContentTypeName select contentType)
                .FirstOrDefault();
        }

        public void Dispose()
        {
            _context.Dispose();
        }
        #endregion

        #region Properties
        public string Name { get; set; }
        public string ContentTypeName { get; set; }
        public int TemplateType { get; set; }
        #endregion
    }
}
