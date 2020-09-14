using System;
using System.Text;

namespace CreateSPSite
{
    class Program
    {
        static void Main()
        {
            Console.OutputEncoding = Encoding.UTF8;
            bool over = false;

            #region Main
            App game = new App(over);

            game.Run();
            #endregion
        }
    }
}
