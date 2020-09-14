using Microsoft.SharePoint.Client;

namespace CreateSPSite.Models
{
    public abstract class AbstractList
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
            var employeesList = _context.Web.Lists.GetByTitle("Employees");
            _context.Load(employeesList);
            _context.ExecuteQuery();
            return employeesList;
        }

        public List CheckExists(ListCollection collection)
        {
            return collection.GetByTitle(Name);
        }
        #endregion

        #region Properties
        public string Name { get; set; }
        public string ContentTypeName { get; set; }
        #endregion
    }
}
