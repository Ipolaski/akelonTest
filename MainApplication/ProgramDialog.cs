namespace MainApplication
{
    public static class ProgramDialog
    {
        public static void Main()
        {
            string directory = string.Empty;
            var userInput = string.Empty;
            while (userInput != "exit" && directory != "exit")
            {
                MenuDialog menu = new();
                directory = menu.GetDirectoryDialog();

                userInput = menu.GetWorkWithFileDialog( directory );

            }
        }
    }
}