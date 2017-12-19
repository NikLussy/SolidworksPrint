using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWX_KKS
{
    class Log
    {
        public static string GetHtmlCode()
        {
            return Head()
                    + Change("V1.2.1.12",
                "Bug Fix",
                new List<string>() {
                    "Ein fehler der die ausführung des Programms verhindert hat, wurde entfernt. Eine unnötige Abfrage (IsToolBoxPart) führte zu diesen fehler. ",
                })
                + Change("V1.2.1.X",
                "Medienberührend",
                new List<string>() {
                    "Die funktion Zertifiziert wir bei einer Zeichnung immer gedruckt, wenn diese angewählt wird.",
                })
                + Change("V1.2.0.X",
                "Es wurden folgende Funktionen hinzugefügt:",
                new List<string>() {
                    "Ein Reiter für die Rüstliste wurde hinzugefügt.",
                })

                + Change("V1.1.0.X",
                "Es wurden folgende Funktionen hinzugefügt:",
                new List<string>() {
                    "Die Zeichnungen mit 'hängenden' Massen werden nicht abgelegt.",
                    "Der Reiter Meldungen wird automatisch gesetzt, wenn ein Eintrag vorhanden ist.",
                    "Dateiname für die Exportdatei wird vorgeschlagen."
                })

                + Change("V1.0.1.X", 
                "Es wurden folgende Funktionen hinzugefügt:", 
                new List<string>() {
                    "Nur freigegebene Zeichnungen exportieren",
                    "Nur Zeichnungen drucken, wenn noch kein PDF abgelegt ist",
                    "Info Fenster mit Change Log"
                })

                + Change("V1.0.0.0",
                "Initialversion",
                new List<string>())
                + Footer();



        }

        private static string Change(string Version, string Description, List<string> Points)
        {
            return Title(Version) + Desc(Description, Points);            

        }

        private static string Title(string title)
        {
            return "<div class=\"Title padding\"><a>" + title + "</a></div>";
        }

        private static string Desc(string desc, List<string> Points)
        {
            return "<div class=\"Desc padding\">" + desc + " " + Point(Points) + "</div>";
        }

        private static string Point(List<string> Points)
        {
            if (Points.Count >= 1)
            {
                string points = "<ul>";
                foreach (string p in Points)
                {
                    points = points + "<li>" + p +"</li>";
                }
                return points + "</ul>";
            }else
            {
                return "";
            }
        }

        private static string Head()
        {
            List<string> head = new List<string>();
            head.Add("<html><head>");
            head.Add("<meta content = \"de-ch\" http - equiv = \"Content-Language\" />");
            head.Add("<meta content = \"text/html; charset=utf-8\" http - equiv = \"Content-Type\" />");
            head.Add("<style>");
            head.Add("body{font-size: 12px;font-family: Arial, Helvetica, sans-serif;}");
            head.Add(".padding{padding: 5px;}");
            head.Add(".Title{font-style: normal;font-size: 14px;background-color: #999999;width: 100%;min-height: 50px;}");
            head.Add(".Desc{background-color: #CCCCCC;width: 100%;min-height: 50px;}");
            head.Add("li{list-style-type: square;list-style-position: outside ;margin-bottom:5px;margin-left: -15px;}");
            head.Add("</style></head><body>");
   
            string headstring = "";
            foreach (string s in head)
            {
                headstring = headstring + " "+ s;
            }


            return headstring;
        }

        private static string Footer()
        {
            List<string> footer = new List<string>();
            footer.Add("</body></html>");
            string footstring = "";
            foreach (string s in footer)
            {
                footstring = footstring + " " + s;
            }


            return footstring;
        }
    }
}
