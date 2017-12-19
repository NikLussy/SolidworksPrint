using SldWorks;
using SwConst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static SWX_KKS.SWX.SWX_Connector;

namespace SWX_KKS.SWX
{
    public class ExportDrawings
    {
        public string FileNameBG = "";
        public string configBg = "";
        public List<Part> Parts = new List<Part>();
        public List<string> Notes = new List<string>();

        public void OverrideMsg()
        {
            if (Settings.PrintOnlyIfNoPDF && Settings.AskOverridePdf)
            {
                Override ov = new Override();
                ov.Text = "Willst du bestehende Dateien überschreiben?";
                if (ov.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Settings.OverridePdf = true;
                    Settings.AskOverridePdf = ov.AskOverride;
                }
                else
                {
                    Settings.OverridePdf = false;
                    Settings.AskOverridePdf = ov.AskOverride;
                }
            }
        }

        public bool GetFile(bool FindFault = false, bool DeleteFault = false)
        {
            try
            {
                if (SWX.Settings.Run)
                {
                   
                    GetActualDoc(new string[] { "modeldoc" });
                    int Mod = modeldoc.GetType();
                    if (FindFault && modeldoc.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                    {
                        GetActualDoc(new string[] { "configuration", "iComponent" });
                        Part part = new Part();

                        modeldoc.ForceRebuild3(true);
                        //Meldung Wiederaufbau fehler wurde gelöscht in absprache mit Pefä                     
                        drawingDoc = (DrawingDoc)modeldoc;
                        IsDangling(FindFault, DeleteFault);
                    }
                    else if (modeldoc != null && modeldoc.GetType() == (int)SwConst.swDocumentTypes_e.swDocASSEMBLY)
                    {

                        GetActualDoc(new string[] { "assemblyDoc", "configuration", "iComponent" });

                        //Lädt die Baugruppe Komplett
                        assemblyDoc.ResolveAllLightWeightComponents(false);

                        FileNameBG = modeldoc.GetTitle();
                        configBg = configuration.Name;
                        return true;
                        //Auswerten();
                    }
                    else if (modeldoc != null && modeldoc.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                    {
                        GetActualDoc(new string[] { "configuration", "iComponent" });

                        int Longstatus = 0;
                        int LongWarnings = 0;
                        Object SheetProps;
                        Part part = new Part();


                        modeldoc.ForceRebuild3(true);
                        //Meldung Wiederaufbau fehler wurde gelöscht in absprache mit Pefä                     
                        drawingDoc = (DrawingDoc)modeldoc;
                        if (!IsDangling())
                        {
                            sheet = (Sheet)drawingDoc.GetCurrentSheet();

                            string FileName = Path.GetFileNameWithoutExtension(modeldoc.GetPathName());
                            part.Name = FileName.Substring(FileName.LastIndexOf("/") + 1);
                            part.PathName = modeldoc.GetPathName();
                            //part.RefConfig = iComponent.ReferencedConfiguration;
                            //part.Released = modeldoc.GetCustomInfoValue("", "Status");
                            part.IsFastener = modeldoc.GetCustomInfoValue("", "isfastener");
                            //Medienberührend funktioniert nicht mit Zeichnung
                            //if (modeldoc.GetCustomInfoValue("", "Medienberührend").ToUpper() == "JA")
                            //    part.Zertifikat = true;
                            //part.Revision = modeldoc.GetCustomInfoValue(part.RefConfig, "Revision");
                            part.Revision = modeldoc.GetCustomInfoValue("", "Revision");
                            modelDocExtension = modeldoc.Extension;
                            //int ret = modelDocExtension.ToolboxPartType;
                            //if (ret != 0)
                            //{
                            //    part.IsToolBoxPart = true;
                            //}

                            if (sheet != null)
                            {
                                SheetProps = sheet.GetProperties();
                                double[] xSheetProps = (double[])SheetProps;
                                //Auftragsnummer
                                SetNote(xSheetProps[5] - 0.015, 0.018, Settings.Auftrag, 5);
                                //Stückzahlen
                                SetNote(xSheetProps[5] - 0.015, xSheetProps[6] - 0.019, (part.Cnt * Settings.Anzahl) + "Stk.", 5);
                                //Kunde
                                SetNote(xSheetProps[5] - 0.015, xSheetProps[6] - 0.005, Settings.Kunde, 4);
                                if (Settings.Zertifiziert)//&& part.Zertifikat
                                {
                                    //Zertifiziert
                                    SetNote(xSheetProps[5] - 0.015, xSheetProps[6] - 0.012, "Zertifiziert", 4);
                                }
                            }

                            string PfdPath = Settings.Path + "\\" + part.Name + "-" + part.Revision + ".pdf";
                            modeldoc.Extension.SaveAs(PfdPath, 0, 1, null, ref Longstatus, ref LongWarnings);

                            if (Settings.DXF)
                            {
                                string DxfPathExt = Settings.Path + "\\" + part.Name + "-" + part.Revision + ".dxf";
                                modeldoc.Extension.SaveAs(DxfPathExt, 0, 1, null, ref Longstatus, ref LongWarnings);
                            }
                            //if (cbZertifiziert.Checked)
                            //{
                            //    StepPathExt = ExportPath + "\\" + ActDoc.Name + "-" + ActDoc.Revision + ".step"; ;
                            //    Model.SaveAs4(StepPathExt, 0, 1, ref Longstatus, ref LongWarnings);
                            //}

                            if (Settings.Print)
                            {
                                Print();
                            }
                            foreach (string note in Notes)
                            {
                                modeldoc.Extension.SelectByID2(note + "@" + sheet.GetName(), "NOTE", 0, 0, 0, false, 0, null, 0);
                                //modeldoc.Extension.SelectAll();
                                modeldoc.EditDelete();
                            }
                        }
                        else
                        {
                            Settings.LogAdd("Fehler in der Zeichnung", modeldoc.GetPathName());
                        }
                    }
                    else
                    {
                        KKS.KKS_Message.Show("Bitte öffne eine Baugruppe oder eine Zeichnung");
                    }
                  
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

        public void GetParts(IComponent2 Comp)
        {
            try
            {
                if (Settings.Run)
                {
                    bool Include = true;
                    Object Children;

                    //TreferencedDoc ActDoc = new TreferencedDoc();

                    if (!Comp.IsSuppressed())//Wenn Zeichnung nicht unterdrückt
                    {

                        Part part = new Part();
                        modeldoc = (ModelDoc2)Comp.GetModelDoc2();
                        if (modeldoc != null)
                        {
                            string FileName = Path.GetFileNameWithoutExtension(modeldoc.GetPathName());
                            part.Name = FileName.Substring(FileName.LastIndexOf("/") + 1);


                            foreach (Part partfromList in Parts)
                            {
                                if (SWX.Settings.NotInBil && Comp.ExcludeFromBOM)
                                {
                                    Include = false;
                                    break;
                                }
                                if (Settings.OnlyAssembly && modeldoc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
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
                                part.PathName = modeldoc.GetPathName();
                                part.RefConfig = Comp.ReferencedConfiguration;
                                part.Released = modeldoc.GetCustomInfoValue("", "Status");
                                part.Beschreibung = modeldoc.GetCustomInfoValue("", "Description");
                                part.IsFastener = modeldoc.GetCustomInfoValue("", "isfastener");
                                if (modeldoc.GetCustomInfoValue("", "Medienberührend").ToUpper() == "JA")
                                    part.Zertifikat = true;
                                part.Revision = modeldoc.GetCustomInfoValue(part.RefConfig, "Revision");
                                modelDocExtension = modeldoc.Extension;
                                //int ret = modelDocExtension.ToolboxPartType;
                                //if (ret != 0)
                                //{
                                //    part.IsToolBoxPart = true;
                                //}
                            }
                        }
                        else
                        {
                            part.PathName = modeldoc.GetPathName();
                            string TempName = Path.GetFileNameWithoutExtension(Comp.GetPathName());
                            part.Name = TempName.Substring(TempName.LastIndexOf("/") + 1);
                            //KKS.KKS_Message.Show(ActDoc.Name);
                            part.RefConfig = "";
                            part.Released = modeldoc.GetCustomInfoValue("", "Status");
                            part.Revision = modeldoc.GetCustomInfoValue(part.RefConfig, "Revision");
                        }
                        if (Include)
                        {
                            Parts.Add(part);
                        }
                    }

                    Children = Comp.GetChildren();
                    foreach (Object com in (Object[])Children)
                    {
                        Child = (IComponent2)com;
                        GetParts(Child);
                    }
                }
            }
            catch (Exception ex)
            {
                if (KKS.KKS_Message.Show("Folgender Fehler ist aufgetreten. Um den Vorgang weiter zu führen Klicke auf OK. /n" + ex.Message, "Fehler", true, "OK", "Abbrechen") == System.Windows.Forms.DialogResult.Cancel)
                {
                    Settings.Run = false;
                }
            }
        }

        private List<string> StatusFreigegeben = new List<string>() {"Freigegeben", "Bibliothek gültig" };

        private List<string> VisibleState = new List<string>() { "Sichtbarkeit unbekannt", "Sichtbar","Halb Sichtbar","Unsichtbar" };

        public bool CheckDrawingRelease()
        {
            if (Settings.Dokfreigabe && !Settings.PrintOnlyReleased)
            {
                for (int x = 0; x < Parts.Count; x++)
                {
                    if (!StatusFreigegeben.Contains(Parts[x].Released))
                    {
                        Settings.NotReleased.Add(Parts[x]);
                    }
                }
                if (Settings.NotReleased.Count >= 1)
                {
                    if (KKS.KKS_Message.Show(Settings.NotReleased.Count.ToString() + " Zeichnungen sind noch nicht freigegeben." + System.Environment.NewLine +
                        "Willst du trotzdem weiterfahren?", "Fehlende Freigaben", true, "Ja", "Nein") == System.Windows.Forms.DialogResult.Cancel)
                    {
                        Settings.Run = false;
                        Settings.Writereleased = true;
                        return false;
                    }
                }
            }
            return true;
        }

        private bool IsDangling(bool FindFault = false,bool DeleteFault = false)
        {
            //Loop für die Sheet's hier einfügen
            swView = drawingDoc.GetFirstView();
            int Fault = 1;
            while (swView != null)
            {
                SldWorks.Annotation swAnot = swView.GetFirstAnnotation3();
                while (swAnot != null)
                {
                    if (swAnot.IsDangling() && (int)swAnot.GetType() == 4 && swAnot.Visible !=3)
                    {
                        if (FindFault)
                        {
                            //swAnot.Visible = (int)swAnnotationVisibilityState_e.swAnnotationVisible;
                            string Position = "";
                            double[] pos = (double[])swAnot.GetPosition();

                            foreach (double p in pos)
                            {
                                Position = Position + Math.Round(p, 4).ToString() + " ";
                            }
                            swAnot.Select3(false, null);
                            modeldoc.ViewZoomToSelection();
                         
                            KKS.KKS_Message.Show("Dein " + Fault + ". Fehler wurde gefunden und herangezoomt." + System.Environment.NewLine 
                                + swAnot.GetName() + System.Environment.NewLine
                                + "Sichtbarkeit: " + VisibleState[swAnot.Visible] + System.Environment.NewLine
                                + "@Position: " + Position);
                            Fault++;
                            if(DeleteFault)
                                modeldoc.EditDelete();
                        }
                        else { return true; }

                    }
                    //else if ((int)swAnot.GetType() == 4)
                    //{
                    //    //string Position = "";
                    //    //int Visible = swAnot.Visible;
                    //    //System.Object Data = swAnot.GetDisplayData();
                    //    //double[] pos = (double[])swAnot.GetPosition();
                    //    //foreach (double p in pos)
                    //    //{
                    //    //    Position = Position + Math.Round(p, 4).ToString() + " ";
                    //    //}
                    //    //string Name = swAnot.GetName();
                    //    swAnot.Select3(false, null);
                    //    modeldoc.ViewZoomToSelection();
                    //}
                    swAnot = swAnot.GetNext3();
                }
                swView = swView.GetNextView();
            }
            return false;
        }

        public void GetDrawing(Part part)
        {
            try
            {
                if (Settings.Run)
                {
                    if (!Settings.PrintOnlyReleased || (Settings.PrintOnlyReleased && StatusFreigegeben.Contains(part.Released)))
                    {
                        int CloseErrors = 0;
                        int CloseWarnings = 0;
                        int Longstatus = 0;
                        int LongWarnings = 0;
                        Object SheetProps;

                        string DwgPath = Path.GetDirectoryName(part.PathName) + "\\" + Path.GetFileNameWithoutExtension(part.PathName) + ".SLDDRW";

                        modeldoc = App.OpenDoc6(DwgPath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref CloseErrors, ref CloseWarnings);
                        if (modeldoc != null)
                        {
                            modeldoc.ForceRebuild3(true);
                            //Meldung Wiederaufbau fehler wurde gelöscht in absprache mit Pefä                     

                            if (modeldoc.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                            {
                                drawingDoc = (DrawingDoc)modeldoc;
                                if (!IsDangling())
                                {
                                    sheet = (Sheet)drawingDoc.GetCurrentSheet();
                                    swRevTable = sheet.RevisionTable;
                                    if (swRevTable != null)
                                    {
                                        string RevTableRevision = swRevTable.GetRevisionForId(1);
                                        if (part.Revision != RevTableRevision && string.IsNullOrEmpty(RevTableRevision))
                                        {
                                            KKS.KKS_Message.Show("Revision der Konfig = " + part.Revision + " Revision der Zeichnung = " + RevTableRevision);
                                        }
                                    }

                                    if (sheet != null)
                                    {
                                        SheetProps = sheet.GetProperties();
                                        double[] xSheetProps = (double[])SheetProps;
                                        //Auftragsnummer
                                        SetNote(xSheetProps[5] - 0.015, 0.018, Settings.Auftrag, 5);
                                        //Stückzahlen
                                        SetNote(xSheetProps[5] - 0.015, xSheetProps[6] - 0.019, (part.Cnt * Settings.Anzahl) + "Stk.", 5);
                                        //Kunde
                                        SetNote(xSheetProps[5] - 0.015, xSheetProps[6] - 0.005, Settings.Kunde, 4);
                                        if (Settings.Zertifiziert && part.Zertifikat)
                                        {
                                            //Zertifiziert
                                            SetNote(xSheetProps[5] - 0.015, xSheetProps[6] - 0.012, "Zertifiziert", 4);
                                        }
                                    }

                                    bool ExportPdf = true;
                                    string PdfPath = Settings.Path + "\\" + part.Name + "-" + part.Revision + " " + KKS.Tools.ReplaceInvalid(part.Beschreibung)  + ".pdf";
                                    if (File.Exists(PdfPath) && Settings.PrintOnlyIfNoPDF)
                                    {
                                        part.Print = false;
                                        if (Settings.AskOverridePdf)
                                        {
                                            Override ov = new Override();
                                            ov.Text = "Die Zeichnung " + part.Name + "-" + part.Revision + ".pdf existiert bereits." + "Willst du die Datei überschreiben?";
                                            if (ov.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                            {
                                                Settings.OverridePdf = true;
                                                Settings.AskOverridePdf = ov.AskOverride;
                                                ExportPdf = true;
                                            }
                                            else
                                            {
                                                Settings.OverridePdf = false;
                                                Settings.AskOverridePdf = ov.AskOverride;
                                                ExportPdf = false;
                                            }
                                        }
                                        else
                                        {
                                            ExportPdf = Settings.OverridePdf;
                                        }
                                    }
                                    if (ExportPdf)
                                    { modeldoc.Extension.SaveAs(PdfPath, 0, 1, null, ref Longstatus, ref LongWarnings); }

                                    if (Settings.DXF)
                                    {
                                        string DxfPathExt = Settings.Path + "\\" + part.Name + "-" + part.Revision + ".dxf";
                                        modeldoc.Extension.SaveAs(DxfPathExt, 0, 1, null, ref Longstatus, ref LongWarnings);
                                    }
                                    //if (cbZertifiziert.Checked)
                                    //{
                                    //    StepPathExt = ExportPath + "\\" + ActDoc.Name + "-" + ActDoc.Revision + ".step"; ;
                                    //    Model.SaveAs4(StepPathExt, 0, 1, ref Longstatus, ref LongWarnings);
                                    //}

                                    if (Settings.Print && part.Print)
                                    {
                                        Print();
                                    }
                                }
                                else
                                {
                                    Settings.LogAdd("Fehler in der Zeichnung", DwgPath, part.Released);
                                }
                                App.CloseDoc(modeldoc.GetPathName());
                                modeldoc = null;
                            }
                        }
                        else
                        {
                            Settings.LogAdd("Nicht gefunden", DwgPath, part.Released);
                        }
                    }else if (!StatusFreigegeben.Contains(part.Released))
                    {
                        Settings.LogAdd("Zeichnung nicht Freigegeben", part.PathName, part.Released);
                    }
                }
            }
            catch (Exception ex)
            {
                if (KKS.KKS_Message.Show("Folgender Fehler ist aufgetreten. Um den Vorgang weiter zu führen Klicke auf OK. /n" + ex.Message, "Fehler", true, "OK", "Abbrechen") == System.Windows.Forms.DialogResult.Cancel)
                {
                    Settings.Run = false;
                }
            }
        }

        private void Print()
        {
            Object vSheetProps;

            drawingDoc = (DrawingDoc)App.ActiveDoc;
            sheet = (Sheet)drawingDoc.GetCurrentSheet();

            PageSetup myPageSetup = null;
            myPageSetup = ((PageSetup)(modeldoc.PageSetup));

            vSheetProps = sheet.GetProperties();
            double[] xSheetProps = (double[])vSheetProps;

            modeldoc.Extension.UsePageSetup = ((int)(swPageSetupInUse_e.swPageSetupInUse_Document));
            //A1=10, A2=9, A3=8, A4Vertical=7,A4=6
            if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA4size)
            {
                modeldoc.Printer = Properties.Settings.Default.DruckerA4;
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A4Fit;
                myPageSetup.ScaleToFit = Fit;
                if(!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A4Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A4Size;
                int Source = Properties.Settings.Default.A4Source;
                if(Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA4sizeVertical)
            {
                modeldoc.Printer = Properties.Settings.Default.DruckerA4;
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A4vFit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A4vScale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A4vSize;
                int Source = Properties.Settings.Default.A4vSource;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA3size)
            {
                modeldoc.Printer = Properties.Settings.Default.DruckerA3;
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A3Fit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A3Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A3Size;
                int Source = Properties.Settings.Default.A3Source;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA2size)
            {
                modeldoc.Printer = Properties.Settings.Default.DruckerA2;
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A2Fit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A2Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A2Size;
                int Source = Properties.Settings.Default.A2Source;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA1size)
            {
                modeldoc.Printer = Properties.Settings.Default.DruckerA1;
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A1Fit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A1Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A1Size;
                int Source = Properties.Settings.Default.A1Source;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
            }
            else
            {
                KKS.KKS_Message.Show("Papierformat nicht im Makro Definiert.");
            }
            modeldoc.PrintDirect();
            //swModel.Extension.PrintOut2(vPageArray, copies, collate, swPrinter, "");
        }

        private void Print2()
        {
            IModelDoc2 swModel;
            PageSetup myPageSetup;
            Object vSheetProps;
            string swPrinter = "";

            drawingDoc = (DrawingDoc)App.ActiveDoc;
            swModel = (ModelDoc2)App.ActiveDoc;
            sheet = (Sheet)drawingDoc.GetCurrentSheet();
            myPageSetup = swModel.Extension.AppPageSetup;
            //myPageSetup = (PageSetup)swModel.PageSetup;
            int pageArray = 0;
            Object vPageArray = pageArray;
            int copies = 1;
            bool collate = true;

            vSheetProps = sheet.GetProperties();
            double[] xSheetProps = (double[])vSheetProps;
            //A1=10, A2=9, A3=8, A4Vertical=7,A4=6
            if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA4size)
            {
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A4Fit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A4Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A4Size;
                int Source = Properties.Settings.Default.A4Source;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
                swPrinter = Properties.Settings.Default.DruckerA4;

                //myPageSetup.DrawingColor = 3;               //Farben in Zeichnungen, Automatisch = 1 Farb-/Grauskalierung = 2 Schwarz und Weiß = 3
                //myPageSetup.ScaleToFit = false;            //Maßstab an Papier anpassen, False/True
                //myPageSetup.Scale2 = 100;           //nicht Auskommentiert, da myPageSetup.ScaleToFit = False
                //myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;          //Ausrichtung Hochformat = 1 Querformat = 2
                //myPageSetup.HighQuality = false;         //Hohe Qualität, False/True
                //myPageSetup.PrinterPaperSize = 9;        //Drucker Papiergröße, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                ////myPageSetup.PrinterPaperSource = 7; //267       'Drucker Papierquelle, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                //swPrinter = Properties.Settings.Default.DruckerA4;
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA4sizeVertical)
            {
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A4vFit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A4vScale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A4vSize;
                int Source = Properties.Settings.Default.A4vSource;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
                swPrinter = Properties.Settings.Default.DruckerA4;
                //myPageSetup.DrawingColor = 3;               //Farben in Zeichnungen, Automatisch = 1 Farb-/Grauskalierung = 2 Schwarz und Weiß = 3
                //myPageSetup.ScaleToFit = false;            //Maßstab an Papier anpassen, False/True
                //myPageSetup.Scale2 = 100;           //nicht Auskommentiert, da myPageSetup.ScaleToFit = False
                //myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;          //Ausrichtung Hochformat = 1 Querformat = 2
                //myPageSetup.HighQuality = false;         //Hohe Qualität, False/True
                //myPageSetup.PrinterPaperSize = 9;        //Drucker Papiergröße, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                ////myPageSetup.PrinterPaperSource = 7; //267       'Drucker Papierquelle, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                //swPrinter = Properties.Settings.Default.DruckerA4;
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA3size)
            {
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A3Fit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A3Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A3Size;
                int Source = Properties.Settings.Default.A3Source;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
                swPrinter = Properties.Settings.Default.DruckerA3;
                //swModel.Extension.AppPageSetup.DrawingColor = 3;               //Farben in Zeichnungen, Automatisch = 1 Farb-/Grauskalierung = 2 Schwarz und Weiß = 3
                //swModel.Extension.AppPageSetup.ScaleToFit = false;            //Maßstab an Papier anpassen, False/True
                //swModel.Extension.AppPageSetup.Scale2 = 100;           //nicht Auskommentiert, da myPageSetup.ScaleToFit = False
                //swModel.Extension.AppPageSetup.Orientation = 2;//(int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;          //Ausrichtung Hochformat = 1 Querformat = 2
                //swModel.Extension.AppPageSetup.HighQuality = false;         //Hohe Qualität, False/True
                //swModel.Extension.AppPageSetup.PrinterPaperSize = 8;        //Drucker Papiergröße, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                ////myPageSetup.PrinterPaperSource = 7; //267       'Drucker Papierquelle, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                //swPrinter = Properties.Settings.Default.DruckerA3;
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA2size)
            {
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A2Fit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A2Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A2Size;
                int Source = Properties.Settings.Default.A2Source;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
                swPrinter = Properties.Settings.Default.DruckerA2;
                //myPageSetup.DrawingColor = 3;             //Farben in Zeichnungen, Automatisch = 1 Farb-/Grauskalierung = 2 Schwarz und Weiß = 3
                //myPageSetup.ScaleToFit = true;            //Maßstab an Papier anpassen, False/True
                ////myPageSetup.Scale2 = 100;           //nicht Auskommentiert, da myPageSetup.ScaleToFit = False
                //myPageSetup.Orientation = 2;//(int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;          //Ausrichtung Hochformat = 1 Querformat = 2
                //myPageSetup.HighQuality = false;         //Hohe Qualität, False/True
                //myPageSetup.PrinterPaperSize = 8;        //Drucker Papiergröße, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                ////myPageSetup.PrinterPaperSource = 7; //267       'Drucker Papierquelle, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                //swPrinter = Properties.Settings.Default.DruckerA2;
            }
            else if (xSheetProps[0] == (int)swDwgPaperSizes_e.swDwgPaperA1size)
            {
                myPageSetup.DrawingColor = 3;
                bool Fit = Properties.Settings.Default.A1Fit;
                myPageSetup.ScaleToFit = Fit;
                if (!Fit)
                {
                    myPageSetup.Scale2 = Properties.Settings.Default.A1Scale;
                }
                myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                myPageSetup.HighQuality = false;
                myPageSetup.PrinterPaperSize = Properties.Settings.Default.A1Size;
                int Source = Properties.Settings.Default.A1Source;
                if (Source != -1)
                {
                    myPageSetup.PrinterPaperSource = Source;
                }
                swPrinter = Properties.Settings.Default.DruckerA1;
                //myPageSetup.DrawingColor = 3;             //Farben in Zeichnungen, Automatisch = 1 Farb-/Grauskalierung = 2 Schwarz und Weiß = 3
                //myPageSetup.ScaleToFit = true;            //Maßstab an Papier anpassen, False/True
                ////myPageSetup.Scale2 = 100;           //nicht Auskommentiert, da myPageSetup.ScaleToFit = False
                //myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;          //Ausrichtung Hochformat = 1 Querformat = 2
                //myPageSetup.HighQuality = false;         //Hohe Qualität, False/True
                //myPageSetup.PrinterPaperSize = 8;        //Drucker Papiergröße, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                ////myPageSetup.PrinterPaperSource = 7; //267       'Drucker Papierquelle, um den Wert zu erhalten ein Makro aufzeichnen und die Druckeinstellungen vornehmen
                //swPrinter = Properties.Settings.Default.DruckerA1;
            }
            else
            {
                KKS.KKS_Message.Show("Papierformat nicht im Makro Definiert.");
            }


            swModel.Extension.PrintOut2(vPageArray, copies, collate, swPrinter, "");

        }

        private void SetNote(double XVal, double YVal, string Text, int Size)
        {
            ModelDoc swModel = (ModelDoc)App.ActiveDoc;
            //TextFormat txtFormat = new TextFormat();

            Note note = (Note)swModel.InsertNote("<FONT color=0x000000ff><FONT size=" + Size + ">" + Text);
            Notes.Add(note.GetName());
            if (note != null)
            {
                note.Angle = 0;
                note.SetTextJustification((int)swTextJustification_e.swTextJustificationRight);
                bool BoolStatus = note.SetBalloon(0, 0);
                Annotation annotation = (Annotation)note.GetAnnotation();
                if (annotation != null)
                {
                    long Longstatus = annotation.SetLeader2(false, 1, true, false, false, false);
                    BoolStatus = annotation.SetPosition(XVal, YVal, 0);
                    //BoolStatus = annotation.SetTextFormat(0, true, txtFormat);
                }
            }
            swModel.ClearSelection();
            swModel.WindowRedraw();
        }

    }
}
