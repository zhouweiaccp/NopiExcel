using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;

namespace ConsoleApplication4
{
    //另外附下源码中的注释部分，关于HSSFDataFormat参数的

    //      0, "General"
    //       1, "0"
    //       2, "0.00"
    //       3, "#,##0"
    //       4, "#,##0.00"
    //       5, "($#,##0_);($#,##0)"
    //       6, "($#,##0_);[Red]($#,##0)"
    //       7, "($#,##0.00);($#,##0.00)"
    //       8, "($#,##0.00_);[Red]($#,##0.00)"
    //       9, "0%"
    //       0xa, "0.00%"
    //       0xb, "0.00E+00"
    //       0xc, "# ?/?"
    //       0xd, "# ??/??"
    //       0xe, "m/d/yy"
    //       0xf, "d-mmm-yy"
    //       0x10, "d-mmm"
    //       0x11, "mmm-yy"
    //       0x12, "h:mm AM/PM"
    //       0x13, "h:mm:ss AM/PM"
    //       0x14, "h:mm"
    //       0x15, "h:mm:ss"
    //       0x16, "m/d/yy h:mm"
   
    //        0x17 - 0x24 reserved for international and Undocumented
    //       0x25, "(#,##0_);(#,##0)"
    //       0x26, "(#,##0_);[Red](#,##0)"
    //       0x27, "(#,##0.00_);(#,##0.00)"
    //       0x28, "(#,##0.00_);[Red](#,##0.00)"
    //       0x29, "_(///#,##0_);_(///(#,##0);_(/// \"-\"_);_(@_)"
    //       0x2a, "_($///#,##0_);_($///(#,##0);_($/// \"-\"_);_(@_)"
    //       0x2b, "_(///#,##0.00_);_(///(#,##0.00);_(///\"-\"??_);_(@_)"
    //       0x2c, "_($///#,##0.00_);_($///(#,##0.00);_($///\"-\"??_);_(@_)"
    //       0x2d, "mm:ss"
    //       0x2e, "[h]:mm:ss"
    //       0x2f, "mm:ss.0"
    //       0x30, "##0.0E+0"
    //       0x31, "@" - This Is text format.
    //       0x31  "text" - Alias for "@"
    class Class2
    {
        public static void testc()
        {

                        //1.创建EXCEL中的Workbook           
            IWorkbook myworkbook = new XSSFWorkbook();  
  
            //2.创建Workbook中的Sheet          
            ISheet mysheet = myworkbook.CreateSheet("sheet1");  
            mysheet.SetColumnWidth(0, 40 * 256);  
  
            //3.创建Row中的Cell并赋值  
            IRow row0 = mysheet.CreateRow(0); row0.CreateCell(0).SetCellValue("130925199662080044");  
            IRow row1 = mysheet.CreateRow(1); row1.CreateCell(0).SetCellValue(""+DateTime.Now+"");  
  
            //4.创建CellStyle与DataFormat并加载格式样式  
            IDataFormat dataformat = myworkbook.CreateDataFormat();  
  
            //【Tips】  
            // 1.使用@ 或 text 都可以  
            // 2.再也不用为身份证号发愁了  
      
            ICellStyle style0 = myworkbook.CreateCellStyle();  
            style0.DataFormat = dataformat.GetFormat("@");  
  
            ICellStyle style1 = myworkbook.CreateCellStyle();  
            style1.DataFormat = dataformat.GetFormat("text");  
         
            //5.将CellStyle应用于具体单元格  
            row0.GetCell(0).CellStyle = style0;  
            row1.GetCell(0).CellStyle = style1;  
         
            //6.保存         
            FileStream file = new FileStream(@"c:\myworkbook9.xlsx", FileMode.Create);  
            myworkbook.Write(file);  
            file.Close();  
        }

    }
}
