using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
 
using Microsoft.Office.Tools.Ribbon;
 
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn2
{
    class ExcelToCsv_Converter
    {
        const string cert = ",";
        const string nextLine = "\r\n";

        class CommandCsv
        {
            public static string Comment    = "//";
            public static string StartTable = "<table>";
            public static string NoExport = "x";
            public static string NewLineConvert = "<br>";// \n or \r\n
           // public static string TabConvert = "<tab>";// \t
            public static string CommaConvert = "<comma>";// ,
            public static string QuotationMarksConvert = "<qm>";// '
            public static string DoubleQuotationMarksConvert = "<dqm>";// "
        }

        static public string ConvertCsv(Excel.Worksheet excelSheet, int endColumn, int endRow, bool commandCheck = true)
        {
            return ConvertCsv(excelSheet, 1, 1, endColumn, endRow, commandCheck);
        }

        static public string ConvertCsv(Excel.Worksheet excelSheet, int beginColumn, int beginRow, int endColumn, int endRow, bool commandCheck = true)
        {
            int startCommandColumn      = -1;
            bool usingNewLineConvert    = true;
           // bool usingTabConvert        = true;
            bool usingCommaConvert      = true;
            bool usingQuotationMarksConvert = true;
            bool usingDoubleQuotationMarksConvert = true;
            if(commandCheck)
            {
                CommandParseInfo commandParseInfo = ParseCommand(excelSheet, beginColumn, beginRow, endColumn, endRow);
                beginColumn     = commandParseInfo.beginColumn;
                beginRow        = commandParseInfo.beginRow;
                endColumn       = commandParseInfo.endColumn;
                endRow          = commandParseInfo.endRow;
                startCommandColumn = commandParseInfo.startCommandColumn;
            }


            StringBuilder csvStringBuild = new StringBuilder();
            for (int j = beginRow; j <= endRow; ++j)
            {
                //check no export row
                if(startCommandColumn > 0)
                {
                    Excel.Range cellValueCancle = excelSheet.Cells[j, startCommandColumn] as Excel.Range;

                    if (cellValueCancle != null && cellValueCancle.Value != null)
                    {
                        string cellValueCancleText = cellValueCancle.Text;
                        cellValueCancleText = cellValueCancleText.ToLower();
                        if (cellValueCancleText == CommandCsv.NoExport)
                        {
                            continue;
                        }
                    }
                }
                 
                for (int i = beginColumn; i <= endColumn; ++i)
                {
                    Excel.Range cellValue = excelSheet.Cells[j, i] as Excel.Range;

                    if (cellValue == null || cellValue.Value == null)
                    {
                        if (i != endColumn)
                        {
                            csvStringBuild.Append(cert);
                        } 
                        continue;
                    }
                    else
                    {
                        string csvCell = cellValue.Text;
                        if(usingNewLineConvert)
                        {
                           csvCell =  csvCell.Replace("\n", CommandCsv.NewLineConvert);
                           csvCell =  csvCell.Replace("\r\n", CommandCsv.NewLineConvert);
                        }
                        //if (usingTabConvert)
                        //{
                        //    csvCell = csvCell.Replace("\t", CommandCsv.TabConvert); 
                        //}
                        if (usingCommaConvert)
                        {
                            csvCell = csvCell.Replace(",", CommandCsv.CommaConvert); 
                        }
                        if(usingQuotationMarksConvert)
                        {
                            csvCell = csvCell.Replace("'", CommandCsv.QuotationMarksConvert); 
                        }
                        if(usingDoubleQuotationMarksConvert)
                        {
                            char []DoubleQuotationMarks = new char[]{'"'};
                            string DoubleQuotationMarksString = new string(DoubleQuotationMarks);
                            csvCell = csvCell.Replace(DoubleQuotationMarksString, CommandCsv.DoubleQuotationMarksConvert); 
                        }

                        csvStringBuild.Append(csvCell);
                        if (i != endColumn)
                        {
                             csvStringBuild.Append(cert);
                        }
                       
                    }
                }
                if (j != endRow)
                {
                    csvStringBuild.Append(nextLine); 
                } 
            }
            return csvStringBuild.ToString();
        }
 

        static public bool ConvertCsvAndSave(Excel.Worksheet excelSheet, int beginColume, int beginRow, int endColume, int endRow, string filePath)
        {
            try
            {
                string csvString = ConvertCsv(excelSheet, beginColume, beginRow, endColume, endRow);
                System.IO.TextWriter l_TextWriter = new System.IO.StreamWriter(filePath, false, Encoding.UTF8);
                l_TextWriter.Write(csvString);
                l_TextWriter.Close();
            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
                return false;
            }
            return true;
        }
        public class CommandParseInfo
        {
            public int beginColumn;
            public int beginRow;
            public int endColumn;
            public int endRow;
            public int startCommandColumn = -1;
        }

        //find Command 
        static public CommandParseInfo ParseCommand(Excel.Worksheet excelSheet, int beginColume, int beginRow, int endColume, int endRow)
        {
            CommandParseInfo commandParseInfo = new CommandParseInfo();
            commandParseInfo.beginColumn    = beginColume;
            commandParseInfo.beginRow       = beginRow;
            commandParseInfo.endColumn      = endColume;
            commandParseInfo.endRow         = endRow;
            commandParseInfo.startCommandColumn = -1;
            bool findStartCommand = false;
            for (int j = beginRow; j <= endRow; ++j)
            {
                for (int i = beginColume; i <= endColume; ++i)
                {
                    Excel.Range cellValue = excelSheet.Cells[j, i] as Excel.Range;

                    if (cellValue == null || cellValue.Value == null)
                    { 
                        continue;
                    }
                    else
                    {
                        string cellString = cellValue.Text;
                        string cellStringLower = cellString.ToLower();
                        if (cellStringLower == CommandCsv.StartTable)
                        {
                            findStartCommand = true;
                            commandParseInfo.startCommandColumn = i;
                            commandParseInfo.beginColumn = i+1;
                            commandParseInfo.beginRow    = j;
                            commandParseInfo.endRow = endRow;
                            for (int k = commandParseInfo.beginColumn;  k <= endColume; ++k)
                             {
                                 Excel.Range cellTableEnd = excelSheet.Cells[j, k] as Excel.Range;
                                 if (cellValue == null || cellValue.Value == null || cellValue.Text == "")
                                {
                                    commandParseInfo.endColumn = k - 1;
                                    break;
                                } 
                             }
                        }
                    }

                    if(findStartCommand == true)
                    {
                        break;
                    }
                }
                if (findStartCommand == true)
                {
                    break;
                }
            }

            return commandParseInfo;
        }
    }
}
