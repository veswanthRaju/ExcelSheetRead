using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace InventoyReport
{
    public class Inventory
    {
        /// <summary>
        /// Provide the required columns data from Excel
        /// </summary>
        /// <param name="fileName">Full path of the excel of format .xlsx</param>
        /// <param name="sheetName">Name of the sheet fo which you want to read the data in Excel</param>
        /// <param name="colKey">Column1, which you want to read from Excel, It will be the key for output data</param>
        /// <param name="colValue">Column2, which you want to read from Excel, It will be the value for output data</param>
        /// <returns>Dictionary with key as #<paramref name="colKey"/> and value as #<paramref name="colValue"/></returns>
        public static Dictionary<string, string> ReadExcel(string fileName, string sheetName, string colKey, string colValue)
        {
            string conString = string.Empty;
            DataTable dtexcel = new DataTable();

            if (fileName.Substring(fileName.LastIndexOf(".")).CompareTo(".xlsx") != 0)
            {
                // conString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007              
                Console.WriteLine("Expected format is .xlsx");
                return null;
            }
            else
                conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HRD=Yes';"; //for above excel 2007                                    

            using (OleDbConnection Connection = new OleDbConnection(conString))
            {
                try
                {
                    Connection.Open();
                    dtexcel = Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //If you apply below commented region below lines should go down..
                    var query = String.Format("select * from [{0}]", sheetName + "$");
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter(query, Connection); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable 
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception occurred ", ex);
                    return null;
                }
                finally
                {
                    Connection.Close();                    
                }
            }

            Dictionary<string, string> dataMap = new Dictionary<string, string>();
            foreach (DataRow dtRow in dtexcel.Rows)
            {
                dataMap.Add(dtRow[colKey].ToString(), dtRow[colValue].ToString());               
            }
            
            return dataMap;

            #region To read more Sheets in single excel..
            //var itemsOfWorksheet = new List<SelectListItem>();
            //if (dtexcel != null)
            //{
            //    //validate worksheet name.
            //    string worksheetName;
            //    for (int cnt = 0; cnt < dtexcel.Rows.Count; cnt++)
            //    {
            //        worksheetName = dtexcel.Rows[cnt]["TABLE_NAME"].ToString();

            //        if (worksheetName.Contains('\''))
            //        {
            //            worksheetName = worksheetName.Replace('\'', ' ').Trim();
            //        }
            //        if (worksheetName.Trim().EndsWith("$"))
            //            itemsOfWorksheet.Add(new SelectListItem { Text = worksheetName.TrimEnd('$'), Value = worksheetName });
            //    }
            //}
            #endregion
        }
    }
}
