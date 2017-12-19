using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using static SWX_KKS.SWX.SWX_Connector;

namespace SWX_KKS
{
    public partial class Main : KKS.F_Vorlage
    {
        #region GUI

        public static int Cnt = 0;

        public Main()
        {
            InitializeComponent();

       
        }

        private void btnPath_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.RestoreDirectory = true;
            sfd.FileName = "Folder";
            sfd.CheckFileExists = false;
            sfd.CheckPathExists = true;
            sfd.ValidateNames = false;
            sfd.Filter = "|(No File)|All Files(*.*)|*.*";
            sfd.FilterIndex = 0;

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                tbSave.Text = System.IO.Path.GetDirectoryName(sfd.FileName);
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                cbA1.Items.Add(printer);
                cbA2.Items.Add(printer);
                cbA3.Items.Add(printer);
                cbA4.Items.Add(printer);
            }
            cbA1.Text = Properties.Settings.Default.DruckerA1;
            cbA2.Text = Properties.Settings.Default.DruckerA2;
            cbA3.Text = Properties.Settings.Default.DruckerA3;
            cbA4.Text = Properties.Settings.Default.DruckerA4;
            tbAuftrag.Text = Properties.Settings.Default.Auftrag;
            tbKunde.Text = Properties.Settings.Default.Kunde;
            tbSave.Text = Properties.Settings.Default.Speicherort;

            LoadSettings();

            webBrowser1.DocumentText = SWX_KKS.Log.GetHtmlCode();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (Printer_Validating())
            {
                Properties.Settings.Default.DruckerA1 = cbA1.Text;
                Properties.Settings.Default.DruckerA2 = cbA2.Text;
                Properties.Settings.Default.DruckerA3 = cbA3.Text;
                Properties.Settings.Default.DruckerA4 = cbA4.Text;
                Properties.Settings.Default.Save();
                SavePaperSetting();
                lblStatus.Text = "Einstellungen gespeichert " + DateTime.Now.ToString();
            }
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.Auftrag = tbAuftrag.Text;
            Properties.Settings.Default.Kunde = tbKunde.Text;
            Properties.Settings.Default.Speicherort = tbSave.Text;
            Properties.Settings.Default.Save();
        }

        private bool Settings_Validating()
        {
            //bool Ok = false;
            List<bool> Ok = new List<bool>();
            Ok.Add(ValidatePath(tbSave));
            Ok.Add(ValidateString(tbKunde));
            Ok.Add(ValidateString(tbAuftrag));

            if (Ok.Contains(false))
                return false;
            else
                return true;
        }

        private bool Printer_Validating()
        {
            List<bool> Ok = new List<bool>();
            Ok.Add(ValidatePrinter(cbA1));
            Ok.Add(ValidatePrinter(cbA2));
            Ok.Add(ValidatePrinter(cbA3));
            Ok.Add(ValidatePrinter(cbA4));

            if (Ok.Contains(false))
                return false;
            else
                return true;
        }

        private bool ValidatePath(TextBox tb)
        {
            bool bStatus = true;
            if (!Directory.Exists(tb.Text)) //.Substring(0,tbPathUsb.Text.Length)
            {
                error.SetError(tb, "Der eingegebene Pfad ist ungültig");
                bStatus = false;
            }
            else
            { error.SetError(tb, ""); }

            return bStatus;
        }

        private bool ValidateString(TextBox tb)
        {
            bool bStatus = true;
            if (string.IsNullOrWhiteSpace(tb.Text))
            {
                error.SetError(tb, "Bitte gib einen Wert ein.");
                bStatus = false;
            }
            else
            { error.SetError(tb, ""); }

            return bStatus;
        }

        private bool ValidatePrinter(ComboBox tb)
        {
            bool bStatus = false;
            try
            {
                foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
                {
                    if (printer.Equals(tb.Text))
                        bStatus = true;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                if (!bStatus) //.Substring(0,tbPathUsb.Text.Length)
                {
                    error.SetError(tb, "Der eingegebene Drucker ist nicht vorhanden");
                }
                else
                { error.SetError(tb, ""); }
            }
            return bStatus;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.RestoreDirectory = true;
            sfd.Filter = "CSV files (*.csv)|*.csv";
            sfd.FilterIndex = 1;
            sfd.FileName = DateTime.Now.ToString("yyyy.MM.dd") + "_Fehlerbericht_SWX_App";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                StreamWriter sWriter = new StreamWriter(sfd.FileName);
                sWriter.WriteLine("Info;Datei;Status");
                foreach (string line in lbMdg.Items)
                {
                    sWriter.WriteLine(line);
                }
                sWriter.Close();
                sWriter.Dispose();
                try { Process.Start(sfd.FileName); } catch (Exception ex) { KKS.KKS_Message.Show(ex.Message); }
                lblStatus.Text = "Speichern Abgeschlossen";
            }
        }

        private void cbFit_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb = (CheckBox)sender;
            switch (cb.Name)
            {
                case "cbA4Fit":
                    lblA4Massstab.Enabled = !cb.Checked;
                    nupA4Scale.Enabled = !cb.Checked;
                    break;
                case "cbA4vFit":
                    lblA4vMassstab.Enabled = !cb.Checked;
                    nupA4vScale.Enabled = !cb.Checked;
                    break;
                case "cbA3Fit":
                    lblA3Massstab.Enabled = !cb.Checked;
                    nupA3Scale.Enabled = !cb.Checked;
                    break;
                case "cbA2Fit":
                    lblA2Massstab.Enabled = !cb.Checked;
                    nupA2Scale.Enabled = !cb.Checked;
                    break;
                case "cbA1Fit":
                    lblA1Massstab.Enabled = !cb.Checked;
                    nupA1Scale.Enabled = !cb.Checked;
                    break;
                case "cbDocProp":
                    cbPrintOnlyReleased.Enabled = cb.Checked;
                    cbPrintIfNoPDF.Enabled = cb.Checked;
                    break;
            }
        }

        private void LoadSettings()
        {
            #region Fit
            cbA1Fit.Checked = Properties.Settings.Default.A1Fit;
            cbA2Fit.Checked = Properties.Settings.Default.A2Fit;
            cbA3Fit.Checked = Properties.Settings.Default.A3Fit;
            cbA4Fit.Checked = Properties.Settings.Default.A4Fit;
            cbA4vFit.Checked = Properties.Settings.Default.A4vFit;
            #endregion

            #region Scale
            nupA1Scale.Value = Properties.Settings.Default.A1Scale;
            nupA2Scale.Value = Properties.Settings.Default.A2Scale;
            nupA3Scale.Value = Properties.Settings.Default.A3Scale;
            nupA4Scale.Value = Properties.Settings.Default.A4Scale;
            nupA4vScale.Value = Properties.Settings.Default.A4vScale;
            #endregion

            #region Size
            foreach (string Key in PaperSize.Keys)
            {
                cbA1Size.Items.Add(Key);
                cbA2Size.Items.Add(Key);
                cbA3Size.Items.Add(Key);
                cbA4Size.Items.Add(Key);
                cbA4vSize.Items.Add(Key);
            }
            cbA1Size.Text = PaperSize.FirstOrDefault(x => x.Value == Properties.Settings.Default.A1Size).Key;
            cbA2Size.Text = PaperSize.FirstOrDefault(x => x.Value == Properties.Settings.Default.A2Size).Key;
            cbA3Size.Text = PaperSize.FirstOrDefault(x => x.Value == Properties.Settings.Default.A3Size).Key;
            cbA4Size.Text = PaperSize.FirstOrDefault(x => x.Value == Properties.Settings.Default.A4Size).Key;
            cbA4vSize.Text = PaperSize.FirstOrDefault(x => x.Value == Properties.Settings.Default.A4vSize).Key;
            #endregion

            #region Source
            foreach (string Key in PaperSource.Keys)
            {
                cbA1Source.Items.Add(Key);
                cbA2Source.Items.Add(Key);
                cbA3Source.Items.Add(Key);
                cbA4Source.Items.Add(Key);
                cbA4vSource.Items.Add(Key);
            }
            cbA1Source.Text = PaperSource.FirstOrDefault(x => x.Value == Properties.Settings.Default.A1Source).Key;
            cbA2Source.Text = PaperSource.FirstOrDefault(x => x.Value == Properties.Settings.Default.A2Source).Key;
            cbA3Source.Text = PaperSource.FirstOrDefault(x => x.Value == Properties.Settings.Default.A3Source).Key;
            cbA4Source.Text = PaperSource.FirstOrDefault(x => x.Value == Properties.Settings.Default.A4Source).Key;
            cbA4vSource.Text = PaperSource.FirstOrDefault(x => x.Value == Properties.Settings.Default.A4vSource).Key;
            #endregion
        }

        private void SavePaperSetting()
        {
            Properties.Settings.Default.A1Fit = cbA1Fit.Checked;
            Properties.Settings.Default.A2Fit = cbA2Fit.Checked;
            Properties.Settings.Default.A3Fit = cbA3Fit.Checked;
            Properties.Settings.Default.A4Fit = cbA4Fit.Checked;
            Properties.Settings.Default.A4vFit = cbA4vFit.Checked;

            Properties.Settings.Default.A1Scale = (int)nupA1Scale.Value;
            Properties.Settings.Default.A2Scale = (int)nupA2Scale.Value;
            Properties.Settings.Default.A3Scale = (int)nupA3Scale.Value;
            Properties.Settings.Default.A4Scale = (int)nupA4Scale.Value;
            Properties.Settings.Default.A4vScale = (int)nupA4vScale.Value;

            Properties.Settings.Default.A1Size = PaperSize[cbA1Size.Text];
            Properties.Settings.Default.A2Size = PaperSize[cbA2Size.Text];
            Properties.Settings.Default.A3Size = PaperSize[cbA3Size.Text];
            Properties.Settings.Default.A4Size = PaperSize[cbA4Size.Text];
            Properties.Settings.Default.A4vSize = PaperSize[cbA4vSize.Text];

            Properties.Settings.Default.A1Source = PaperSource[cbA1Source.Text];
            Properties.Settings.Default.A2Source = PaperSource[cbA2Source.Text];
            Properties.Settings.Default.A3Source = PaperSource[cbA3Source.Text];
            Properties.Settings.Default.A4Source = PaperSource[cbA4Source.Text];
            Properties.Settings.Default.A4vSource = PaperSource[cbA4vSource.Text];
            Properties.Settings.Default.Save();
        }

        private Dictionary< string,int> PaperSize = new Dictionary<string, int>() {
            {"A4",9},
            {"A3",8},
            {"A2",7},
            {"A1",6}};

        private Dictionary<string, int> PaperSource = new Dictionary<string, int>() {
            {"Default",-1},
            {"Auto",7},
            {"Kassete 1",14},
            {"Kassete 2",3},
            {"Kassete 3",2},
            {"Kassete 4",257}};
        #endregion

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (Settings_Validating())
            {
                SaveSettings();

                SWX.Settings.Run = true;
                SWX.Settings.NotInBil = cbNotInBil.Checked;
                SWX.Settings.Print = cbPrint.Checked;
                SWX.Settings.Dokfreigabe = cbDocProp.Checked;
                if (SWX.Settings.Dokfreigabe)
                {
                    SWX.Settings.PrintOnlyIfNoPDF = cbPrintIfNoPDF.Checked;
                    SWX.Settings.PrintOnlyReleased = cbPrintOnlyReleased.Checked;
                }
                else
                {
                    SWX.Settings.PrintOnlyIfNoPDF = false;
                    SWX.Settings.PrintOnlyReleased = false;
                }
                SWX.Settings.Zertifiziert = cbZertifiziert.Checked;
                SWX.Settings.DXF = cbCreateDxf.Checked;

                SWX.Settings.OnlyAssembly = cbOnlyAssembly.Checked;

                SWX.Settings.Auftrag = tbAuftrag.Text;
                SWX.Settings.Kunde = tbKunde.Text;
                SWX.Settings.Path = tbSave.Text;
                SWX.Settings.Anzahl = (int)nupAnz.Value;

                SWX.Settings.Log.Clear();
                SWX.Settings.NotReleased.Clear();

                lbMdg.Items.Clear();
                btnStart.Enabled = false;
                btnCancel.Enabled = true;
                progressBar1.Visible = true;

                bgWorker.RunWorkerAsync();

            }
        }

        private void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                SWX.SWX_Connector.GetApp();
                SWX.ExportDrawings Exp = new SWX.ExportDrawings();
                Exp.OverrideMsg();
                if (!Exp.GetFile())
                    return;

                //Check if Run OK
                if (bgWorker.CancellationPending || SWX.Settings.Run == false)
                { RunCancel(ref e); return; }
                else
                    bgWorker.ReportProgress(3);

                Exp.GetParts(SWX.SWX_Connector.iComponent);

                //Check if Run OK
                if (bgWorker.CancellationPending || SWX.Settings.Run == false)
                { RunCancel(ref e); return; }
                else
                    bgWorker.ReportProgress(10);

                if (!Exp.CheckDrawingRelease())
                { RunCancel(ref e); return; }

                //Berechnung der einzelnen Schritte der Progressbar
                decimal Procent = 80 / Exp.Parts.Count;

                for (int x = 0; x < Exp.Parts.Count; x++)
                {
                    bgWorker.ReportProgress(25 + (int)Math.Round(x * Procent, 0));
                    Exp.GetDrawing(Exp.Parts[x]);

                    if (bgWorker.CancellationPending || SWX.Settings.Run == false)
                    { RunCancel(ref e); return; }
                }
                return;
            }

            catch (Exception ex)
            {
                KKS.KKS_Message.Show(ex.Message);
                return;
            }
        }

        private void RunCancel(ref DoWorkEventArgs e)
        {
            e.Cancel = true;
            bgWorker.ReportProgress(0);
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (e.Cancelled)
                {
                    lblStatus.Text = "Abbruch durch User";
                }

                // Check to see if an error occurred in the background process.

                else if (e.Error != null)
                {
                    lblStatus.Text = "Ein Fehler ist aufgetreten. Bitte Fragen sie Ihren Artzt oder Apotheker";
                }
                else
                {
                    // Everything completed normally.
                    lblStatus.Text = "Auftrag erledigt";
                }

                //Change the status of the buttons on the UI accordingly
                btnStart.Enabled = true;
                btnCancel.Enabled = false;
                progressBar1.Visible = false;
                SWX.Settings.Run = false;
                if (SWX.Settings.Writereleased)
                {
                    for (int x = 0; x < SWX.Settings.NotReleased.Count; x++)
                    {
                       lbMdg.Items.Add(SWX.Settings.NotReleased[x].Name + "Zeichnung nicht Freigegeben;" + SWX.Settings.NotReleased[x].PathName + ";" + SWX.Settings.NotReleased[x].Released);
                    }
                    SWX.Settings.Writereleased = false;
                    tabControl1.SelectedTab = tpMdg;
                }
                else
                {
                    if (SWX.Settings.Log.Count < 1)
                    {
                        lbMdg.Items.Add("Keine Meldung");
                    }
                    else
                    {
                        for (int x = 0; x < SWX.Settings.Log.Count; x++)
                        {
                            lbMdg.Items.Add(SWX.Settings.Log[x]);
                        }
                        tabControl1.SelectedTab = tpMdg;
                    }
                }
            }
            catch (Exception ex)
            {
                KKS.KKS_Message.Show(ex.Message);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            SWX.Settings.Run = false;
            if (bgWorker.IsBusy)
            {
                bgWorker.CancelAsync();
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            SWX.SWX_Connector.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //SWX.SWX_Connector.GetApp();
            //GetActualDoc(new string[] { "modeldoc" });
            //drawingDoc = (SldWorks.DrawingDoc)modeldoc;
            ////Loop für die Sheet's hier einfügen
            //SldWorks.View swView = drawingDoc.GetFirstView();
            //while (swView != null)
            //{
            //    SldWorks.Annotation swAnot = swView.GetFirstAnnotation3();
            //    while(swAnot != null)
            //    {
            //        if(swAnot.IsDangling())
            //        {
            //            swAnot.Select3(false, null);
            //            modeldoc.ViewZoomToSelection();
            //            goto FoundDangling;
            //        }
            //        swAnot = swAnot.GetNext3();
            //    }
            //    swView = swView.GetNextView();
            //}
            //FoundDangling:

            ////SldWorks.ModelDocExtension ext = SWX.SWX_Connector.modeldoc.Extension.
            //SWX.SWX_Connector.modeldoc.Extension.SelectByID2("RD3" + "@" + "Zeichenansicht1", "DIMENSION", 0, 0, 0, false, 0, null, 0);
            ////modeldoc.Extension.SelectAll();
            ////SWX.SWX_Connector.modeldoc.EditDelete();
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            lblStatus.Text = "";
        }

        private void btnFault_Click(object sender, EventArgs e)
        {
            try
            {
                btnFault.Enabled = false;
                SWX.SWX_Connector.GetApp();
                SWX.Settings.Run = true;
                SWX.ExportDrawings Exp = new SWX.ExportDrawings();
                Exp.GetFile(true);
                SWX.Settings.Run = false;
                btnFault.Enabled = true;
                lblStatus.Text = "Fehlersuche abgeschlossen";
            }
            catch (Exception ex)
            {
                KKS.KKS_Message.Show(ex.Message);
            }
        }

        private void btnFaultDelete_Click(object sender, EventArgs e)
        {
            try
            {
                btnFault.Enabled = false;
                SWX.SWX_Connector.GetApp();
                SWX.Settings.Run = true;
                SWX.ExportDrawings Exp = new SWX.ExportDrawings();
                Exp.GetFile(true);
                SWX.Settings.Run = false;
                btnFault.Enabled = true;
                lblStatus.Text = "Fehlersuche abgeschlossen";
            }
            catch (Exception ex)
            {
                KKS.KKS_Message.Show(ex.Message);
            }
        }

        private void btnRuestliste_Click(object sender, EventArgs e)
        {
            lblRunning.Visible = true;
            SWX.Settings.NotInBil = cbNotInBilExport.Checked;

            SWX.SWX_Connector.GetApp();
            SWX.ExportList Exp = new SWX.ExportList();
            if (!Exp.GetFile())
                return;

            lblCnt.Text = Exp.GetParts(SWX.SWX_Connector.iComponent).ToString() + " Gezählt";

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.RestoreDirectory = true;
            sfd.Filter = "Excel Files| *.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK)
                Excel.WriteExcel(Excel.GetBauteile(Exp.Parts),sfd.FileName);

            lblRunning.Visible = false;
            lblCnt.Visible = true;

            if (File.Exists(sfd.FileName))
                Process.Start(sfd.FileName);

            for (int x = 0; x < SWX.Settings.Log.Count; x++)
            {
                lbMdg.Items.Add(SWX.Settings.Log[x]);
            }
            tabControl1.SelectedTab = tpMdg;
        }
    }
}
