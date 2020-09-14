using CreateSPSite.Models;
using Microsoft.SharePoint.Client;
using System;

namespace CreateSPSite.Factories
{
    public class ContentTypeFactory
    {
        private ClientContext _context;
        private AbstractContentType _model;

        public ContentTypeFactory(ClientContext context)
        {
            _context = context;
        }

        public void GetContentType(string contentTypeName)
        {
            switch(contentTypeName)
            {
                case Constants.ContentType.Employee:
                    _model = new EmployeeContentType(_context);
                    _model.CreateFields(_model.Create());
                    break;
                case Constants.ContentType.Project:
                    _model = new ProjectContentType(_context);
                    _model.CreateFields(_model.Create());
                    break;
                case Constants.ContentType.ProjectDoc:
                    _model = new ProjectDocContentType(_context);
                    _model.CreateFields(_model.Create());
                    break;
                default:
                    throw new Exception("Chương trình Không hỗ trợ content type này");
            }
            _model.Dispose();
        }
    }
}
