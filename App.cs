using System;
using System.Threading;

namespace CreateSPSite
{
    public class App
    {
        private bool _over;
        private ConsoleKeyInfo _key;

        public App(bool over)
        {
            _over = over;
        }

        public void Run()
        {
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
            Console.WriteLine("[1] Create Employees list [2] Create Project list [3] Create Project Document list [4] Create a site and all list [Esc or Ctrl-C] Exit");
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
            SharepointService.CreateEmployeeContentType();
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
