namespace CreateSPSite
{
    class Program
    {
        static void Main()
        {
            bool over = false;

            #region Main
            App game = new App(over);

            game.Run();
            #endregion
        }
    }
}
