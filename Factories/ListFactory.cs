using CreateSPSite.Models;
using Microsoft.SharePoint.Client;
using System;

namespace CreateSPSite.Factories
{
    public class ListFactory
    {
        private ClientContext _context;
        private AbstractList _model;

        public ListFactory(ClientContext context)
        {
            _context = context;
        }

        public void CreateList(string listName)
        {
            switch(listName)
            {
                case Constants.ListTitle.Employees:
                    _model = new EmployeesList(_context);
                    _model.Create();
                    break;
                default:
                    throw new Exception("Chương trình không hỗ trợ List này");
            }
            _model.Dispose();
        }
    }
}
