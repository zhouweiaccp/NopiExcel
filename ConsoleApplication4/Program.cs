using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication4
{

    class info
    {
        public string name { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            //System.Data.DataTable dt = new NetUtilityLib.ExcelHelper("c:\\1.xls").ExcelToDataTable("Sheet1");
            //if (dt != null)
            //{
            //    Console.WriteLine("eee");
            //}
         
            //Class1.DataTableToExcel(dt, "C:\\abc"+DateTime.Now.ToString("MMss")+".xls");
         //   Console.ReadLine();


            Class2.testc();
        }

        public static System.Data.DataTable dtaa()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("id", typeof (int));
            dt.Columns.Add("name", typeof(string));

            //for (int i = 0; i < 10; i++)
            //{
            //    dt.Rows.c
            //}
            return dt;
        }
    }
}
