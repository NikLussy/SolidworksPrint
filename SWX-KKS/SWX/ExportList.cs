using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SldWorks;
using System.IO;
using SwConst;
using System.Diagnostics;
using static SWX_KKS.SWX.SWX_Connector;

namespace SWX_KKS.SWX
{
    class ExportList
    {

        public string FileNameBG = "";
        public string configBg = "";
        public List<Part> Parts = new List<Part>();
        private int Cnt = 0;
        //public List<string> Notes = new List<string>();

        public bool GetFile(bool FindFault = false)
        {
            try
            {

                GetActualDoc(new string[] { "modeldoc" });
                //int Mod = modeldoc.GetType();
                if (modeldoc != null && modeldoc.GetType() == (int)SwConst.swDocumentTypes_e.swDocASSEMBLY)
                {

                    GetActualDoc(new string[] { "assemblyDoc", "configuration", "iComponent" });

                    //Lädt die Baugruppe Komplett
                    assemblyDoc.ResolveAllLightWeightComponents(false);

                    //FileNameBG = modeldoc.GetTitle();
                    //configBg = configuration.Name;
                    Cnt = 0;
                    return true;
                }
                else
                {
                    KKS.KKS_Message.Show("Bitte öffne eine Baugruppe");
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (KKS.KKS_Message.Show("Folgender Fehler ist aufgetreten. Um den Vorgang weiter zu führen Klicke auf OK." + ex.Message, "Fehler", true, "OK", "Abbrechen") == System.Windows.Forms.DialogResult.Cancel)
                {
                    Settings.Run = false;
                }
            }
            return false;
        }
        //public Stopwatch t = new Stopwatch();

        public int GetParts(IComponent2 Comp)
        {
            try
            {
                bool Include = true;
                Object Children;
              
                if (!Comp.IsSuppressed())//Wenn Zeichnung nicht unterdrückt
                {
                    Part part = new Part();
                    modeldoc = (ModelDoc2)Comp.GetModelDoc2();
                    if (modeldoc != null)
                    {
                        //string FileName = Path.GetFileNameWithoutExtension(modeldoc.GetPathName());
                        part.Name = modeldoc.GetCustomInfoValue("", "DokID");
                        //t.Stop();
                        //Settings.LogAdd(t.Elapsed.Milliseconds.ToString(), Cnt.ToString(), part.Name);
                        //t.Reset();
                        //t.Start();
                        foreach (Part partfromList in Parts)
                        {
                            if (SWX.Settings.NotInBil && Comp.ExcludeFromBOM)
                            {
                                Include = false;
                                break;
                            }
                            if (partfromList.Name == part.Name)
                            {
                                partfromList.Cnt++;
                                Include = false;
                                break;
                            }
                        }

                        if (Include)
                        {
                            if (modeldoc.GetCustomInfoValue("", "Medienberührend").ToUpper() == "JA")
                                part.Zertifikat = true;

                            part.ArtNr = modeldoc.GetCustomInfoValue("", "Artikelnummer");
                            part.Beschreibung = modeldoc.GetCustomInfoValue("", "Description");
                            Parts.Add(part);
                        }
                    }
                    if (Include)
                    {
                        
                    }
                }

                Children = Comp.GetChildren();
                Cnt++;
                foreach (Object com in (Object[])Children)
                {
                    Child = (IComponent2)com;
                    GetParts(Child);
                }
            }
            catch (Exception ex)
            {
                if (KKS.KKS_Message.Show("Folgender Fehler ist aufgetreten. Um den Vorgang weiter zu führen Klicke auf OK. /n" + ex.Message, "Fehler", true, "OK", "Abbrechen") == System.Windows.Forms.DialogResult.Cancel)
                {
                   
                }
            }
            return Cnt;
        }
    }
}
