using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.ComponentModel;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Excel;
namespace DrawBallot
{
    static class Methods
    {
        private static Random rng = new Random();
        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
        public static string DataTableToCSV(this DataTable datatable, char seperator)
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataRow dr in datatable.Rows)
            {
                for (int i = 0; i < datatable.Columns.Count; i++)
                {
                    sb.Append(dr[i].ToString());

                    if (i < datatable.Columns.Count - 1)
                        sb.Append(seperator);
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        public static DataTable ConvertCSVtoDataTable1(string strFilePath)
        {
            StreamReader sr = new StreamReader(strFilePath);
            DataTable dt = new DataTable();
            var tempvalue ="";

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("isDrawed", typeof(bool));
            int index = 0;
            while (!sr.EndOfStream)
            {
                tempvalue = sr.ReadLine();
                string[] rows = new string[tempvalue.Split(';').Length];
                index = 0;
                foreach (string value in tempvalue.Split(';'))
                {
                    rows[index] = value;
                    index++;
                }
                
                DataRow dr = dt.NewRow();
                string rowValue = null;
                for (int i = 0; i < rows.Length-1; i++)
                {       
                    rowValue += rows[i];
                    rowValue += '-';
                }
                rowValue = rowValue.Substring(0, rowValue.Length - 1);
                dr[0] = rowValue;
                dr[1] = ToBoolean(rows[rows.Length-1]);
                
                dt.Rows.Add(dr);
            }
            sr.Close();
            return dt;
        }

        public static DataTable ConvertCSVtoDataTable2(string strFilePath)
        {
            StreamReader sr = new StreamReader(strFilePath);
            
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("First Name");
            dt.Columns.Add("Last Name");
            dt.Columns.Add("Prize");
            dt.Columns.Add("Date");
            while (!sr.EndOfStream)
            {
                string[] rows = Regex.Split(sr.ReadLine(), ";");
                DataRow dr = dt.NewRow();
                for (int i = 0; i < 5; i++)
                {
                        dr[i] = rows[i];
                }
                dt.Rows.Add(dr);
            }
            sr.Close();
            return dt;
        }

        public static string ExcelToCSV(string FilePath, char Seperator)
        {
            
            StringBuilder sb = new StringBuilder();
            FileStream stream = File.Open(FilePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet Result = excelReader.AsDataSet();
            excelReader.Close();
            int rowCount = 0;
            while (rowCount < Result.Tables[0].Rows.Count)
            {
                for (int i = 0;i < Result.Tables[0].Columns.Count; i++)
                {
                    sb.Append(Result.Tables[0].Rows[rowCount][i].ToString());
                    sb.Append(';');
                }
                sb.Append("false");
                sb.AppendLine();
                rowCount++;
            }
            return sb.ToString();
        }
        public static bool ToBoolean(this string value)
        {
            switch (value.ToLower())
            {
                case "true":
                    return true;
                case "t":
                    return true;
                case "1":
                    return true;
                case "0":
                    return false;
                case "false":
                    return false;
                case "f":
                    return false;
                default:
                    throw new InvalidCastException("You can't cast a weird value to a bool!");
            }
        }
        public static ListViewItem CreateItem(string[] arr)
        {
            ListViewItem item = new ListViewItem(arr);
            return item;
        }
        public static string ListToCSV(BindingList<string> list, char seperator)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string item in list)
            {
                sb.Append(item);
                sb.Append(seperator);
            }
            sb.Remove(sb.Length - 1, 1);
            return sb.ToString();
        }
    }
    
}
