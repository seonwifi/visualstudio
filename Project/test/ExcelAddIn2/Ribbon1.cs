using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

  
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelAddIn2.ExcelExport;

namespace ExcelAddIn2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
             
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("aaa");
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            //string test = excelSheet.Cells[1, 1].Value.ToString();
            //패스 다이알로그
          // string savePath = Globals.ThisAddIn.Application.GetSaveAsFilename(Type.Missing, "Excel Files (*.xls), *.xls") as string;
            //현재 열여있는 패쓰
            //Globals.ThisAddIn.Application.ActiveWorkbook.Path


            //Excel.Range ran = excelSheet.Range;


           // System.Windows.Forms.MessageBox.excelSheetShow(Globals.ThisAddIn.Application.ActiveWorkbook.Path);
            //excelSheet
            try
            {
                CsClass cc = new CsClass(PublicType.PublicType_public, "myclass", "");

                cc.AddVar(new CsVariable(PublicType.PublicType_public, "int", "m_intvalue", "0"));

                string str = cc.MakeString();


                System.IO.TextWriter l_TextWriter = new System.IO.StreamWriter(@"g:/newclass.cs", false, Encoding.UTF8);
                l_TextWriter.Write(str);
                l_TextWriter.Close();

                if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
                {
                    return;
                }

                Excel.Worksheet excelSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                if (excelSheet == null)
                {
                    return;
                }

                string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
                filePath += "/" + excelSheet.Name + ".csv";
                int endColumn = excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                int endRow = excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                ExcelToCsv_Converter.ConvertCsvAndSave(excelSheet, 1, 1, endColumn, endRow, filePath);




               //System.Windows.Forms.MessageBox.Show(excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row.ToString());
               // System.Windows.Forms.MessageBox.Show(excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column.ToString());

               // System.Windows.Forms.MessageBox.Show(excelSheet.Cells.get_End(Excel.XlDirection.xlUp).Row.ToString());
               // System.Windows.Forms.MessageBox.Show(excelSheet.Cells.get_End(Excel.XlDirection.xlUp).Column.ToString());


                //Excel.Range cellValue = excelSheet.Cells[1, 3] as Excel.Range;
                //if (cellValue != null)
                // {
                //     System.Windows.Forms.MessageBox.Show(cellValue.Value);
                // }
                //if (cellValue == null)
                //{
                //    System.Windows.Forms.MessageBox.Show("value null");
                //}
                //else
                //{
                //    cellValue.AddComment("코멘트 입니다.");
                //    //System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                //    //cellValue.Interior.Color = Excel.XlRgbColor.rgbBlue;
                //    cellValue.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                //    //System.Windows.Forms.MessageBox.Show((cellValue.Text as string));
                //    //if (cellValue.Value == null)
                //    //{
                //    //    System.Windows.Forms.MessageBox.Show(" if (cellValue.Value == null)");
                //    //}
                //    //else
                //    //{
  
                //    //    System.Windows.Forms.MessageBox.Show(cellValue.Value.ToString());
                //    //}

                //    //if (cellValue.Value2 == null)
                //    //{
                //    //    System.Windows.Forms.MessageBox.Show(" if (cellValue.Value2 == null)");
                //    //}
                //    //else
                //    //{

                //    //    System.Windows.Forms.MessageBox.Show(cellValue.Value2.ToString());
                //    //}
                //}
            }
            catch (System.Exception ee)
            {
                System.Windows.Forms.MessageBox.Show(ee.Message);
               // System.Windows.Forms.MessageBox.Show("수정 중인 작업을 끝내 주세요.");
            }
           // System.Windows.Forms.MessageBox.Show(excelSheet.get_Range("A").Count.ToString());
            //System.Windows.Forms.MessageBox.Show(excelSheet.Range.Column.ToString());
           // System.Windows.Forms.MessageBox.Show(excelSheet.Range.Row.ToString()); 
           // System.IO.TextWriter textWriter = new System.IO.StreamWriter("");
        }
    }
}
