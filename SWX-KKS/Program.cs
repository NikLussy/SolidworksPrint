using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SWX_KKS
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (SWX_KKS.Properties.Settings.Default.UpdateSettings)
            {
                SWX_KKS.Properties.Settings.Default.Upgrade();
                SWX_KKS.Properties.Settings.Default.UpdateSettings = false;
                SWX_KKS.Properties.Settings.Default.Save();
            }


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //KKS.KKS_Message.Show("Bitte bezahle deine Lizenzgebühren");
            Application.Run(new Main());
        }
    }
}
