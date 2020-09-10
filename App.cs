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
            Thread.Sleep(2000);
        }

        private void Update()
        {
            Console.WriteLine("Welcome...");
            Console.WriteLine("Please select 1 action");
            Console.WriteLine("[1] Create Employees list [2] Create Project list [3] Create Project Document list [Esc or Ctrl-C] Exit");
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
                    Console.WriteLine("You press 2");
                    break;
                case ConsoleKey.D3:
                    Console.WriteLine("You press 3");
                    break;
                case ConsoleKey.D4:
                    // Delete content type by name
                    SharepointService.DeleteContentType("Employee1");
                    break;
                case ConsoleKey.D5:
                    // Delete content type by name
                    SharepointService.FindContentTypeAssoc("Employee1");
                    break;
                case ConsoleKey.Escape:
                    _over = true;
                    break;
                default:
                    return;
            }
        }
    }
}
