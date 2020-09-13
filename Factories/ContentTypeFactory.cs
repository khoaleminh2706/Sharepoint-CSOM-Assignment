using CreateSPSite.Models;
using Microsoft.SharePoint.Client;
using System;

namespace CreateSPSite.Factories
{
    public class ContentTypeFactory
    {
        private ClientContext _context;
        public ContentTypeFactory(ClientContext context)
        {
            _context = context;
        }

        public void GetContentType(string contentTypeName)
        {
            switch(contentTypeName)
            {
                case Constants.ConteType.Employee:
                    var model = new EmployeeContentType(_context);
                    model.Create();
                    Console.WriteLine("Content Tpe fisnis");
                    break;
                default:
                    throw new Exception("Không hỗ trợ contenttype này");
            }
        }
    }
}
