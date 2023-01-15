using System;
using System.Windows.Forms;

namespace ExcelReaderWriter
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("ODEzODYyQDMyMzAyZTMxMmUzMEVlakJZa1k3UzZ2TlZLajNwZXRIQ0loNE15TiszempnaVkySm5BVTRiRXc9");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}