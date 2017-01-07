using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using AuthorizeNet;

namespace AuthorizeTransactions
{
    class ExcelReport
    {

        static public void FormatAsTable(Excel.Range SourceRange, string TableName, string TableStyleName)
        {
            SourceRange.Worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
            SourceRange, System.Type.Missing, XlYesNoGuess.xlYes, System.Type.Missing).Name =
                TableName;
            SourceRange.Select();
            SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;
        }

        static public bool GenerateReport(List<AuthorizeNet.Transaction> transactions, out string fileName) {

            fileName = "";

            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.DisplayAlerts = false;
                if (xlApp == null)
                {
                    return false;
                }


                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                //Excel.Range xRange;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //HEADER
                xlWorkSheet.Cells[1, 1] = "Response Code";
                xlWorkSheet.Cells[1, 2] = "Submit Date/Time";
                xlWorkSheet.Cells[1, 3] = "Card Number";
                //xlWorkSheet.Cells[1, 4] = "Invoice Number";
                
                xlWorkSheet.Cells[1, 4] = "Invoice Description";
                xlWorkSheet.Cells[1, 5] = "Total Amount";
                xlWorkSheet.Cells[1, 6] = "CC Comm";
                xlWorkSheet.Cells[1, 7] = "Method";
                //xlWorkSheet.Cells[1, 8] = "Action Code";
                xlWorkSheet.Cells[1, 8] = "Customer First Name";
                //xlWorkSheet.Cells[1, 9] = "Company";
                double totalCCComm = 0;
                double amex = 0;
                double visa = 0;
                double mstc = 0;
                double others = 0;
                double total = 0;
                int count = 2;
                foreach (AuthorizeNet.Transaction item in transactions)
                {
                    if (item.ResponseCode == 1)
                    {

                        //Find credit card comm
                        double ccComm = 0;
                        foreach (AuthorizeNet.LineItem li in item.LineItems) {
                            if (li.ID == "CMM") {
                                ccComm = (double)li.UnitPrice;
                                break;
                            }
                        }

                        xlWorkSheet.Cells[count, 1] = item.ResponseCode;
                        xlWorkSheet.Cells[count, 2] = item.DateSubmitted;
                        xlWorkSheet.Cells[count, 3] = item.CardNumber;
                        //xlWorkSheet.Cells[count, 4] = item.InvoiceNumber;
                        xlWorkSheet.Cells[count, 4] = item.OrderDescription;
                        xlWorkSheet.Cells[count, 5] = item.AuthorizationAmount;
                        xlWorkSheet.Cells[count, 6] = ccComm;
                        xlWorkSheet.Cells[count, 7] = item.CardType;
                        //xlWorkSheet.Cells[count, 8] = item.CardResponseCode;
                        xlWorkSheet.Cells[count, 8] = item.FirstName;
                        //xlWorkSheet.Cells[count, 9] = item;

                        total += (double)item.AuthorizationAmount;
                        totalCCComm += ccComm;

                        switch (item.CardType) {
                            case "AmericanExpress":
                                amex += (double)item.AuthorizationAmount;
                                break;
                            case "MasterCard":
                                mstc += (double)item.AuthorizationAmount;
                                break;
                            case "Visa":
                                visa += (double)item.AuthorizationAmount;
                                break;
                            default:
                                others += (double)item.AuthorizationAmount;
                                break;
                        }

                        count++;
                    }
                }

                Excel.Range SourceRange = (Excel.Range)xlWorkSheet.get_Range("A1", "H" + count.ToString()); // or whatever range you want here
                FormatAsTable(SourceRange, "Table1", "TableStyleMedium2");

                count++;
                xlWorkSheet.Cells[count, 4] = "TOTAL";
                xlWorkSheet.Cells[count, 5] = total;
                count++;
                count++;
                xlWorkSheet.Cells[count, 4] = "VISA";
                xlWorkSheet.Cells[count, 5] = visa;
                count++;
                xlWorkSheet.Cells[count, 4] = "MASTERCARD";
                xlWorkSheet.Cells[count, 5] = mstc;
                count++;
                xlWorkSheet.Cells[count, 4] = "AMEX";
                xlWorkSheet.Cells[count, 5] = amex;
                count++;
                xlWorkSheet.Cells[count, 4] = "OTHERS";
                xlWorkSheet.Cells[count, 5] = others;
                count++;
                count++;
                xlWorkSheet.Cells[count, 4] = "TOTAL CC COMM";
                xlWorkSheet.Cells[count, 5] = totalCCComm;


                fileName = "c:\\Reports\\Escrow" + DateTime.Now.ToString("yyyyMMdd") + ".xls";

                //xRange = xlWorkSheet.get_Range("A1", "G" + count.ToString());
                xlWorkSheet.Columns.AutoFit();

                xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                return true;
            }
            catch (Exception es) {
                return false;
            }
            finally { 

            }

        }

        static private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
