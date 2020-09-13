using Microsoft.SharePoint.Client;
using System;

namespace CreateSPSite.Models
{
    public abstract class AbstractField: IDisposable
    {
        protected ClientContext _context;

        public AbstractField(ClientContext context)
        {
            _context = context;
        }

        #region Methods
        public virtual void Create() {}

        protected Field GetField()
        {
            return _context.Web.AvailableFields.GetByInternalNameOrTitle(InternalName);
        }

        public void Dispose()
        {
            _context.Dispose();
        }
        #endregion

        #region Properties
        public ContentType TargetContentType { get; set; }
        public string InternalName { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; } = "New Custom Field";
        public string XmlSchema { get; set; }
        #endregion
    }
}
