using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWX_KKS.SWX
{
    class Settings
    {
        public static bool Run = false;
        public static bool NotInBil = false;
        public static bool Print = false;
        public static bool Dokfreigabe = false;
        public static bool Zertifiziert = false;
        public static bool DXF = false;
        public static bool PrintOnlyIfNoPDF = false;
        public static bool PrintOnlyReleased = false;
        public static bool OnlyAssembly = false;
        public static int Anzahl = 1;
        public static string Auftrag = "";
        public static string Kunde = "";
        public static string Path = "";
        public static bool OverridePdf = false;
        public static bool AskOverridePdf = true;

        public static bool Writereleased = false;
        public static List<string> Log = new List<string>();
        public static List<Part> NotReleased = new List<Part>();


        public static void LogAdd(string Col1, string Col2 = "", string Col3 = "", string Col4 = "", string Col5 = "")
        {
            Log.Add(Col1 + ";" + Col2 + ";" + Col3 + ";" + Col4 + ";" + Col5);
        }
    }
}
