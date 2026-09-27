using System;
using System.Windows.Forms;

namespace Sortix
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize(); // .NET 6+
            // Для .NET Framework 4.x замени на:
            // Application.EnableVisualStyles();
            // Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new Form1());
        }
    }
}