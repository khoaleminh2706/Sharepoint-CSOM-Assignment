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
            Console.WriteLine("[1] Create Employees list");
            Console.WriteLine("[2] Create Project list");
            Console.WriteLine("[3] Create Project Document list");
            Console.WriteLine("[4] Create Site");
            Console.WriteLine("[Esc or Ctrl-C] Exit");
            _key = Console.ReadKey();

            // xuống 1 dòng
            Console.WriteLine();
            HandleKey(_key);
        }

        public void HandleKey(ConsoleKeyInfo key)
        {
            switch (key.Key)
            {
                case ConsoleKey.D1:
                    SharepointService.CreateEmployeeContentType();
                    break;
                case ConsoleKey.D2:
                    SharepointService.CreateProjectList1();
                    break;
                case ConsoleKey.D3:
                    Console.WriteLine("You press 3");
                    // TODO: Delete Project Documents
                    break;
                case ConsoleKey.Escape:
                    _over = true;
                    break;
                default:
                    return;
            }
        }

        private void HandleCreateList(string listName)
        {

        }
    }
}
