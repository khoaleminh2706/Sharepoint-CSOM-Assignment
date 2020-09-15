using System;
using CreateSPSite.Factories;
using CreateSPSite.Provider;
using CreateSPSite.Services;
using Microsoft.SharePoint.Client;

namespace CreateSPSite
{
    public class App
    {
        private bool _over;
        private ConsoleKeyInfo _key;
        private string _loginName;
        private string _password;
        private string _siteUrl;
        private SPClientContextProvider _provider;
        private SharepointService _service;

        private ContentTypeFactory _contentTypeFactory;
        private ListFactory _listFactory;

        public App(bool over)
        {
            _over = over;
        }

        public void Run()
        {
            // TODO: allow changing url
            Console.WriteLine("Welcome...");
            Console.WriteLine("Vui lòng điền thông tin đăng nhập:");
            Console.Write("Login Name: ");
            _loginName = Console.ReadLine();
            Console.Write("Password: ");
            _password = Console.ReadLine();
            Console.Write("Site Url: ");
            _siteUrl = Console.ReadLine();

            _provider = new SPClientContextProvider(_loginName, _password, _siteUrl);
            ResetContext(_siteUrl);

            while (!_over)
            {
                Update(); 
            }
            Console.WriteLine("Exiting...");
        }

        private void Update()
        {
            Console.WriteLine("Please select 1 action");
            Console.WriteLine("[1] Create Employees list");
            Console.WriteLine("[2] Create Project list");
            Console.WriteLine("[3] Create Project Document list");
            Console.WriteLine("[4] Create Site");
            Console.WriteLine("[5] Current Url");
            Console.WriteLine("[6] Change Site Url");

            Console.WriteLine("[Esc or Ctrl-C] Exit");
            _key = Console.ReadKey();

            // xuống 1 dòng
            Console.WriteLine();
            HandleKeyPress(_key);
        }

        public void HandleKeyPress(ConsoleKeyInfo key)
        {
            try
            {
                switch (key.Key)
                {
                    case ConsoleKey.D1:
                        Console.WriteLine("Start creating Employee...");
                        _contentTypeFactory.GetContentType(Constants.ContentType.Employee);
                        AccessHrSite();
                        _listFactory.CreateList(Constants.ListTitle.Employees);
                        Console.WriteLine("Finish creating Employee...");
                        ResetContext(_siteUrl);
                        break;
                    case ConsoleKey.D2:
                        Console.WriteLine("Start creating project...");
                        _contentTypeFactory.GetContentType(Constants.ContentType.Project);
                        AccessHrSite();
                        _listFactory.CreateList(Constants.ListTitle.Projects);
                        Console.WriteLine("Finish creating project...");
                        ResetContext(_siteUrl);
                        break;
                    case ConsoleKey.D3:
                        Console.WriteLine("Start creating project document...");
                        _contentTypeFactory.GetContentType(Constants.ContentType.ProjectDoc);
                        AccessHrSite();
                        _listFactory.CreateList(Constants.ListTitle.ProjDoc);
                        Console.WriteLine("Finish creating project document...");
                        ResetContext(_siteUrl);
                        break;
                    case ConsoleKey.D4:
                        HandleCreateSite();
                        break;
                    case ConsoleKey.D5:
                        Console.WriteLine(_provider.SiteUrl);
                        break;
                    case ConsoleKey.D6:
                        HandleChangeUrl();
                        break;
                    case ConsoleKey.Escape:
                        _over = true;
                        break;
                    default:
                        return;
                }
            }
            catch (Exception ex)
            {
                Console.Write("Lỗi App: ");
                Console.WriteLine(ex.GetType().Name + " " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        private void HandleChangeUrl()
        {
            Console.Write("Nhập link mới: ");
            _siteUrl = Console.ReadLine();
            ResetContext(_siteUrl);
        }

        private void AccessHrSite()
        {
            Web hrWeb;
            try
            {
                hrWeb = _service.CheckHRSubsiteExist();
            }
            catch (Exception)
            {
                Console.WriteLine("subite HR không tồn tại.");
                Console.Write("Bạn có muốn tạo subsite [Y] Có [N] Tạo trên trang này.");
                if (Console.ReadKey().Key == ConsoleKey.Y)
                    hrWeb = _service.CreateHRSubsite();
                else
                    return;
            }
            ResetContext(hrWeb.Url);
        }

        private void HandleCreateSite()
        {
            Console.WriteLine("Create Site");
            Console.Write("Admin site Url: ");
            string adminUrl = Console.ReadLine();
            Console.Write("Root site Url: ");
            string rootSiteUrl = Console.ReadLine();
            Console.Write("Site Title: ");
            string siteTitle = Console.ReadLine();
            Console.Write("Site Url: ");
            string url = Console.ReadLine();

            ResetContext(adminUrl);

            string siteUrl = _service.CreateSite(rootSiteUrl, _loginName, siteTitle, url);
            
            ResetContext(siteUrl);
            _service.CreateHRSubsite();

            _contentTypeFactory.GetContentType(Constants.ContentType.Employee);
            _contentTypeFactory.GetContentType(Constants.ContentType.Project);
            _contentTypeFactory.GetContentType(Constants.ContentType.ProjectDoc);

            AccessHrSite();

            _listFactory.CreateList(Constants.ListTitle.Employees);
            _listFactory.CreateList(Constants.ListTitle.Projects);
            _listFactory.CreateList(Constants.ListTitle.ProjDoc);
        }

        private void ResetContext(string url)
        {
            _provider.SiteUrl = url;
            var context = _provider.Create();
            _contentTypeFactory = new ContentTypeFactory(context);
            _listFactory = new ListFactory(context);
            _service = new SharepointService(context);
        }
    }
}
