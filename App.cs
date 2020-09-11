using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace CreateSPSite
{
    public class App
    {
        private bool _over;
        private ConsoleKeyInfo _key;
        private string loginName;
        private string password;
        private string ITFirm;

        public App(bool over)
        {
            _over = over;
        }

        public void Run()
        {
            Console.WriteLine("");
            while (!_over)
            {
                Update(); 
            }
            Console.WriteLine("Exiting...");
        }

        private void Update()
        {
            Console.WriteLine("Welcome...");
            Console.WriteLine("Please select 1 action");
            Console.WriteLine("[1] Create Employees list");
            Console.WriteLine("[2] Create Project list");
            Console.WriteLine("[3] Create Project Document list");
            Console.WriteLine("[4] Create a site and all list");
            Console.WriteLine("[C] Change Site Url");
            Console.WriteLine("[Esc or Ctrl-C] Exit");
            _key = Console.ReadKey();

            // xuống 1 dòng
            Console.WriteLine();
            HandleKeyPress(_key);
        }

        public void HandleKeyPress(ConsoleKeyInfo key)
        {
            switch (key.Key)
            {
                case ConsoleKey.D1:
                    HandleOption1();
                    break;
                case ConsoleKey.D2:
                    HandleOption2();
                    break;
                case ConsoleKey.D3:
                    HandleOption3();
                    break;
                case ConsoleKey.D4:
                    HandleOption4();
                    break;
                case ConsoleKey.Escape:
                    _over = true;
                    break;
                default:
                    return;
            }
        }

        private void HandleOption1()
        {
            Console.WriteLine("Create Employees list");
            Console.WriteLine("Please insert website url");
            string siteUrl = Console.ReadLine();
            SharepointService.CreateEmployeeContentType(siteUrl);
        }

        private void HandleOption2()
        {
            SharepointService.CreateProjectList();
        }

        private void HandleOption3()
        {
            Console.WriteLine("Create Document list");
            SharepointService.CreateDocumentList();
        }

        private void HandleOption4()
        {
            Console.WriteLine("Create Site");
            string siteUrl = SharepointService.CreateSite("https://khoaleminh-admin.sharepoint.com", "https://khoaleminh.sharepoint.com", "new site 2", "newsite2");
            Console.WriteLine(siteUrl);
        }
    }
}
