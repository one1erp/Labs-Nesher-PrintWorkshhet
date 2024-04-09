using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace PrintWorkshhet
{


    [ComVisible(true)]
    [ProgId("PrintWorkshhet.PrintWorkshhetcls")]
    public class PrintWorkshhetCls : IEntityExtension
    {
        private INautilusServiceProvider sp;

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }

        #region members
        private IDataLayer dal;
        public Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
        public Microsoft.Office.Interop.Excel._Worksheet ExcelWorkSheet;
        private DAL.Worksheet currentWorksheet;
        #endregion

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                #region Init

                sp = Parameters["SERVICE_PROVIDER"];
                var ntlsCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlsCon);
                var records = Parameters["RECORDS"];
                var id = records.Fields["WORKSHEET_ID"].Value.ToString();
                var x = long.Parse(id);
                dal = new DataLayer();
                dal.Connect();
                records = null;
                #endregion


                //Get specified worksheet
                currentWorksheet = dal.GetWorksheetById(x);


                //Get aliquots from worksheet
                var entries = currentWorksheet.WORKSHEET_ENTRY;
                List<Aliquot> aliquots =
                    (from entry in entries where entry.ALIQUOT != null orderby entry.WORKSHEET_ORDER select entry.ALIQUOT).ToList();

                //Set data into object
                var printDetailses = new List<PrintDetails>();

                //Get worksheet tests 
                List<Test> tests = aliquots.SelectMany(aliquot => aliquot.Tests).ToList();

                //Get worksheet results           
                List<Result> results = tests.SelectMany(test => test.Results).ToList();


                //Get result templates by worksheet sessions
                var sessions = currentWorksheet.WORKSHEET_SESSION;
                List<long> resultTemplateIds = new List<long>();
                foreach (var session in sessions)
                {
                    var res = session.WORKSHEET_TEMPLATE_SESSION.WORKSHEET_TEMPLATE_RESULT.ToList();
                    foreach (WORKSHEET_TEMPLATE_RESULT templateResult in res)
                    {
                        resultTemplateIds.Add(templateResult.RESULT_TEMPLATE_ID);
                    }
                }

                foreach (var result in results)
                {

                    var pd = printDetailses.FirstOrDefault(p => p.TestTemplateName == result.Test.TestTemplate.Name);
                    if (pd == null)
                    {
                        printDetailses.Add(new PrintDetails(result));
                    }
                    else
                    {
                        pd.AddResult(result);
                    }
                }


                string[] firstRow = { currentWorksheet.NAME, currentWorksheet.WORKSHEET_TEMPLATE.DESCRIPTION };
                ExportDtataTableToExcel(printDetailses, firstRow);
                SaveFile(currentWorksheet.NAME);
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                MessageBox.Show("Error : אנא פנה לתמיכה" + "\n" + ex.Message);
            }
            finally
            {
                dal.Close();

                ExcelApp.Quit();

                ExcelApp = null;

                ExcelWorkSheet = null;

            }
        }


        public void ExportDtataTableToExcel(List<PrintDetails> list, string[] firstRow)
        {

            int b = 1;


            //Create excel
            var workbook = ExcelApp.Workbooks.Add(Missing.Value);
            ExcelWorkSheet = ExcelApp.ActiveSheet;
            ExcelApp.DefaultSheetDirection = (int)XlDirection.xlToLeft;



            foreach (var pd in list)
            {


                if (b > 2)
                {
                    //Create new sheet 
                    ExcelWorkSheet = (Excel.Worksheet)workbook.Worksheets.Add();
                }
                else
                {
                    ExcelWorkSheet = (Excel.Worksheet)workbook.Worksheets.Item[b];
                }
                b++;


                ExcelWorkSheet.Name = pd.TestTemplateName.MakeSafeFilename('_');

                //Set first row
                for (int i = 0; i < firstRow.Length; i++)
                {
                    ExcelWorkSheet.Cells[1, i + 1] = firstRow[i];
                    ExcelWorkSheet.Cells[1, i + 1].Font.Bold = true;
                }

                //Set A2
                ExcelWorkSheet.Cells[2, 1] = currentWorksheet.CREATED_ON.ToString();
                ExcelWorkSheet.Cells[2, 1].Font.Bold = true;

                //Set A3
                int row = 3;
                ExcelWorkSheet.Cells[row, 1] = "Aliquot Name";
                ExcelWorkSheet.Cells[row, 1].Font.Bold = true;


                //Set columns
                var resultNames = pd.GetResultNames();
                int col;
                for (col = 0; col < resultNames.Count + 0; col++)
                {
                    ExcelWorkSheet.Cells[row, col + 2] = resultNames[col];
                    ExcelWorkSheet.Cells[row, col + 2].Font.Bold = true;
                }


                var lastCol = GetExcelColumnName(col + 1);
                var row3Range = ExcelWorkSheet.Range["B" + row, lastCol + "" + row];
                row3Range.WrapText = true;
                row3Range.EntireColumn.ColumnWidth = 8.5;


                var distinctaliq = pd.Aliquots.Distinct().ToList();
                for (int i = 0; i < distinctaliq.Count; i++)
                {
                    ExcelWorkSheet.Cells[++row, 1] = distinctaliq[i].Name;
                }
                var row1Range = ExcelWorkSheet.Range["A1", "A" + "" + distinctaliq.Count + 2];
                row1Range.EntireColumn.ColumnWidth = 22;
                //CONSTANTS
                int aliquotFirstCol = 1;
                int aliquotFirstRow = 4;
                int aliquotRows = distinctaliq.Count + aliquotFirstRow;
                int resultTitlesRow = 3;
                int resultTitlesFirstCol = 2;
                int resultTitlesColumn = resultNames.Count + 2;


                for (int i = aliquotFirstRow; i < aliquotRows; i++)
                {
                    for (int j = resultTitlesFirstCol; j < resultTitlesColumn; j++)
                    {
                        string aliqName = null;
                        var aliqcell = ExcelWorkSheet.Cells[i, aliquotFirstCol].Value;
                        if (aliqcell != null)
                        {
                            aliqName = aliqcell.ToString();
                            string resultName = null;
                            var resultCell = ExcelWorkSheet.Cells[resultTitlesRow, j].Value;
                            if (resultCell != null)
                            {
                                resultName = resultCell.ToString();
                                if (pd.HasResult(aliqName, resultName))
                                {
                                    ExcelWorkSheet.Cells[i, j].Interior.Color = ColorTranslator.ToOle(Color.Gray);
                                }
                            }
                        }

                    }
                }

                Microsoft.Office.Interop.Excel.Range c1 = ExcelWorkSheet.Cells[1, 1];
                Microsoft.Office.Interop.Excel.Range c2 = ExcelWorkSheet.Cells[row, resultNames.Count + 1];
                var oRange = (Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet.get_Range(c1, c2);
                var a = oRange.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom];
                a.Weight = 2d;

                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders.Color = Color.Black;



            }

        }
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }



        public void SaveFile(string wn)
        {
            var ph = dal.GetPhraseByName("Location folders");
            var pe = ph.PhraseEntries.Where(p => p.PhraseDescription == "Worksheet result entry").FirstOrDefault();
            string fileName = pe.PhraseName + "\\" + wn + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            ExcelWorkSheet.SaveAs(fileName);

            ExcelApp.Quit();

            OpenExcel(fileName);
        }

        private void OpenExcel(string fileName)
        {
            Process.Start(fileName + ".xlsx");
        }
    }
}
