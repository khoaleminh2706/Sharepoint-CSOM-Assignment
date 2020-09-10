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
            Console.WriteLine("Choose one options");
            Console.WriteLine("[1] Create Employees list");
            Console.WriteLine("[2] Create Project list");
            Console.WriteLine("[3] Create Project Document list");
            Console.WriteLine("[Esc] Exit");
            _key = Console.ReadKey();
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
                case ConsoleKey.Escape:
                    _over = true;
                    break;
                default:
                    return;
            }
        }
    }
}
