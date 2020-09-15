using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
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
            var contentTypeColl = _context.Site.RootWeb.ContentTypes;
            _context.Load(_context.Web.Lists);
            _context.Load(contentTypeColl);
            _context.ExecuteQuery();

            var list = CheckListExists(_context.Web.Lists);
           
            if (list != null)
                throw new Exception("List đã tồn tại");

            // Check content type exists
            var targetContentType = GetContentType(contentTypeColl);
            if (targetContentType == null)
                throw new Exception($"Content Type {ContentTypeTitle} không tồn tại. Vui lòng tạo content type trước.");

            ListCreationInformation creationInfo = new ListCreationInformation
            {
                Title = Title,
                Description = "New list description",
                TemplateType = TemplateType
            };

            List newList = _context.Web.Lists.Add(creationInfo);
            _context.Load(newList, li => li.ContentTypes);
            _context.ExecuteQuery();

            newList.ContentTypesEnabled = true;
            newList.ContentTypes.AddExistingContentType(targetContentType);
            
            if (TemplateType == (int)ListTemplateType.GenericList)
            {
                var itemContentType = GetContentType(newList.ContentTypes, "Item");
                if (itemContentType != null)
                    itemContentType.DeleteObject();
            }

            newList.Update();
            _context.ExecuteQuery();

            AddView(newList);
            AddCustomColum(newList);
            _context.ExecuteQuery();

            return newList;
        }

        protected List CheckListExists(ListCollection collection)
        {
            return (from list in collection where list.Title == Title select list)
                .FirstOrDefault();
        }

        protected ContentType GetContentType(ContentTypeCollection collection, string contentTypeTitle = "")
        {
            contentTypeTitle = contentTypeTitle != "" ? contentTypeTitle : ContentTypeTitle;
            return (from contentType in collection where contentType.Name == contentTypeTitle select contentType)
                .FirstOrDefault();
        }

        protected virtual List AddCustomColum(List list) => list;

        protected virtual void AddView(List list)
        {
            // load data
            // load all fields
            _context.Load(list.Fields);

            var targetView = list.Views.GetByTitle(ViewTitle);
            _context.Load(targetView, v => v.ViewFields);
            _context.ExecuteQuery();

            var fields = list.Fields.Where(fi => ColumnList.Contains(fi.InternalName)).ToList();

            fields.ToList().ForEach(fi =>
            {
                targetView.ViewFields.Add(fi.InternalName);
            });
            targetView.Update();
        }

        public void Dispose()
        {
            _context.Dispose();
        }
        #endregion

        #region Properties
        public string Title { get; set; }
        public string ContentTypeTitle { get; set; }
        public int TemplateType { get; set; } = (int)ListTemplateType.GenericList;
        public string ViewTitle { get; set; } = "All Items";
        public List<string> ColumnList { get; set; }
        #endregion
    }
}
