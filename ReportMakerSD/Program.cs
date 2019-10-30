using System;
using System.Windows.Forms;

namespace ReportMakerSD
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static int Main(string[] args)
        {
            Console.WriteLine("Командная строка содержит " + args.Length + " агрумента.");
            Console.WriteLine("Вот они: ");
            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine(args[i]);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }
    }
}
