using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SWX_KKS
{
    class Excel
    {
        public static string[,] GetBauteile(List<SWX.Part> Parts)
        {
            int num = 5;
            int count = Parts.Count;

            string[,] strArray = new string[count+1, num];
        
            strArray[0, 0] = "Art-Nr";
            strArray[0, 1] = "DokID";
            strArray[0, 2] = "Beschreibung";
            strArray[0, 3] = "Anzahl";
            strArray[0, 4] = "Medienberührend";

            int Row = 1;
            foreach (SWX.Part part in Parts)
            {
                strArray[Row, 0] = part.ArtNr;
                strArray[Row, 1] = part.Name;
                strArray[Row, 2] = part.Beschreibung;
                strArray[Row, 3] = part.Cnt.ToString();
                strArray[Row, 4] = part.Zertifikat.ToString();
                //strArray[Row, 5] = ;
                Row++;
            }
            return strArray;
        }

        public static bool WriteExcel(string[,] Values,  string FileName = "", int TableBegin = 0, int Col = 1, int StartRow = 1)
        {
            KKS.xls.CheckExcellProcesses();
            Microsoft.Office.Interop.Excel.Application oExcel = null;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = null;
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = null;

            try
            {
                oExcel = new Microsoft.Office.Interop.Excel.Application();
                xlWorkbook = oExcel.Workbooks.Add();
                xlWorksheet = xlWorkbook.Sheets[1];

                Microsoft.Office.Interop.Excel.Range range = xlWorksheet.get_Range(KKS.Tools.IntToAA(Col) + StartRow, KKS.Tools.IntToAA(Values.GetLength(1) + Col - 1) + (Values.GetLength(0)));
                //string Row = KKS.Tools.IntToAA(Col) + "2:" + KKS.Tools.IntToAA(Values.GetLength(1)+ Col-1) + Values.GetLength(0);
                range.Value2 = Values;

                #region FormatTableAndSheet
                range = xlWorksheet.get_Range(KKS.Tools.IntToAA(Col) + (StartRow + TableBegin), KKS.Tools.IntToAA(Values.GetLength(1) + Col - 1) + (Values.GetLength(0)));
                range.Columns.AutoFit();

                #endregion

              
                if (File.Exists(FileName))
                    File.Delete(FileName);
                xlWorkbook.SaveAs(FileName);
                //xlWorkbook.SaveAs(FileName);
                xlWorkbook.Close(false);
            }
            catch (Exception ex)
            {
                KKS.KKS_Message.Show(ex.Source + ": " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcel);
                KKS.xls.KillExcel();
            }
            return true;
        }
    }
}
