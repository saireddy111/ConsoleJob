using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SplitFilesJob.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace SplitFilesJob
{
    class Program
    {
        static void Main(string[] args)
        {
            string destPath = "D:\\Modified_Files";

            Console.WriteLine("Initiating the process...");
            Console.WriteLine("Reading the File...");
            DataTable dt = GetDataTableFromExcel("D:\\April PBS Statements.xlsx");

            List<ClientModel> mList = 
             (from DataRow dr in dt.Rows
                     select new ClientModel()
                     {
                         Client = dr["Client"].ToString(),
                         GroupNumber = dr["Group Number"].ToString(),
                         AssociationNumber = dr["Association Number"].ToString(),
                         MerchantNumber = dr["Merchant Number"].ToString(),
                         MerchantDBAName = dr["Merchant DBA Name"].ToString(),
                         MerchantPricingMonthEndIDDisplay = dr["Merchant Pricing Month End ID Display"].ToString(),
                         MerchantPricingCategoryDescription = dr["Merchant Pricing Category Description"].ToString(),
                         MerchantPricingFeeItemName_Billed = dr["Merchant Pricing Fee Item Name - Billed"].ToString(),
                         MerchantPricingKey2Description = dr["Merchant Pricing Key 2 Description"].ToString(),
                         MerchantPricingGrossCount = dr["Merchant Pricing Gross Count"].ToString(),
                         MerchantPricingNetAmount = dr["Merchant Pricing Net Amount"].ToString(),
                         MerchantPricingFeePerItemRate = dr["Merchant Pricing Fee Per Item Rate"].ToString(),
                         MerchantPricingFeePercentage = dr["Merchant Pricing Fee Percentage"].ToString(),
                         MerchantPricingInterchangePerItemRate = dr["Merchant Pricing Interchange Per Item Rate"].ToString(),
                         MerchantPricingInterchangePercentRate = dr["Merchant Pricing Interchange Percent Rate"].ToString(),
                         MerchantPricingTotalFees = dr["Merchant Pricing Total Fees"].ToString()

                     }).ToList();

            Console.WriteLine("Transformed the Datatable to List.");

            var distinctClients = mList.Where(x => (x.MerchantNumber ?? string.Empty) != string.Empty)
                                       .Select(x => new { x.MerchantNumber }).Distinct().ToList();

            Console.WriteLine(string.Format("Filtered distinct clients. Found:{0}", distinctClients.Count));

            foreach (var client in distinctClients)
            {
                List<string> distinctMonths = mList.Where(x => x.MerchantNumber == client.MerchantNumber
                                               && (x.MerchantPricingMonthEndIDDisplay ?? string.Empty) != string.Empty)
                                            .Select(x => x.MerchantPricingMonthEndIDDisplay).Distinct().ToList();
                Console.WriteLine(string.Format("Filtered distinct Months. Found:{0}", distinctMonths.Count));

                for (int i = 0; i < distinctMonths.Count; i++)
                {
                    List<ClientModel> filteredList = mList.Where(x => x.MerchantNumber == client.MerchantNumber
                    && (x.MerchantPricingMonthEndIDDisplay == distinctMonths.ElementAt(i) || x.MerchantPricingMonthEndIDDisplay == string.Empty)).ToList();

                    Console.WriteLine(string.Format("Filtered client rows w.r.t Month. Found:{0} rows", filteredList.Count));

                    string filePath = Path.Combine(destPath, string.Format("{0}-{1}-{2}.xlsx", client.MerchantNumber, "MonthlyStatement", distinctMonths.ElementAt(i)));

                    Console.WriteLine(string.Format("Writing the filtered data for {0} to the file {1}", distinctMonths.ElementAt(i), filePath));
                    SaveDataSetAsExcel(ListToDataTable(filteredList), dt.Columns, filePath);
                }
            }
        }

        public static DataTable ListToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            // Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties((BindingFlags.Public | BindingFlags.Instance));
            foreach (PropertyInfo prop in Props)
            {
                // Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }

            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            // put a breakpoint here and check datatable
            return dataTable;
        }

        public static void SaveDataSetAsExcel(DataTable dataTable, DataColumnCollection Columns, string exceloutFilePath)
        {
            using (var fs = new FileStream(exceloutFilePath, FileMode.Append, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet(dataTable.TableName);
                List<string> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in Columns)
                {
                    //columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                columnIndex = 0;
                foreach (System.Data.DataColumn column in dataTable.Columns)
                {
                    columns.Add(column.ColumnName);                   
                    columnIndex++;
                }

                int rowIndex = 1;
                XSSFCellStyle stylePercentage = null;
                XSSFCellStyle styleCurrency = null;
                foreach (DataRow dsrow in dataTable.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        var cell = row.CreateCell(cellIndex);
                        
                        if (col == "MerchantPricingFeePercentage" || col == "MerchantPricingInterchangePercentRate")
                        {
                            stylePercentage = (XSSFCellStyle)workbook.CreateCellStyle();
                            stylePercentage.DataFormat = workbook.CreateDataFormat().GetFormat("0.00%");                           

                            cell.SetCellType(CellType.Numeric);
                            cell.CellStyle = stylePercentage;
                            if (dsrow[col].ToString() != "")
                            {
                                cell.SetCellValue(Convert.ToDouble(dsrow[col].ToString()));
                            }                            
                        }
                        else if (col == "MerchantPricingNetAmount" 
                            || col == "MerchantPricingFeePerItemRate" 
                            || col == "MerchantPricingInterchangePerItemRate"
                            || col == "MerchantPricingTotalFees")
                        {
                            styleCurrency = (XSSFCellStyle)workbook.CreateCellStyle();
                            styleCurrency.DataFormat = workbook.CreateDataFormat().GetFormat("$#,##0.00");
                            cell.SetCellType(CellType.Numeric);

                            cell.CellStyle = styleCurrency;
                            if (dsrow[col].ToString() != "")
                            {
                                cell.SetCellValue(Convert.ToDouble(dsrow[col].ToString()));
                            }
                        }
                        else if (col == "GroupNumber"
                            || col == "AssociationNumber"
                            || col == "MerchantNumber"
                            || col == "MerchantPricingMonthEndIDDisplay"
                            || col == "MerchantPricingGrossCount")
                        {
                            cell.SetCellType(CellType.Numeric);
                            if (dsrow[col].ToString() != "")
                            {
                                cell.SetCellValue(Convert.ToDouble(dsrow[col].ToString()));
                            }
                        }
                        else
                        {
                            cell.SetCellValue(dsrow[col].ToString());
                        }
                        
                        excelSheet.AutoSizeColumn(cellIndex);
                        cellIndex++;                        
                    }
                    rowIndex++;
                }
                workbook.Write(fs);
            }
        }

        private static DataTable GetDataTableFromExcel(String Path)
        {
            XSSFWorkbook wb;
            ISheet sheet;
            String Sheet_name;
            using (var fs = new FileStream(Path, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);
                Sheet_name = wb.GetSheetAt(0).SheetName;  //get first sheet name
            }

            sheet = wb.GetSheet(Sheet_name);
            wb.Close();

            DataTable dt = new DataTable(sheet.SheetName);

            /*IRow headerRow = sheet.GetRow(1);
            foreach (ICell headerCell in headerRow)
            {
                if (headerCell.ToString() != "")
                {
                    dt.Columns.Add(headerCell.ToString());
                }
            }*/

            int rowIndex = 0;
            foreach (IRow row in sheet)
            {              

                if (row.FirstCellNum == 0)
                {
                    if (dt.Columns.Count > 0 && row.GetCell(0).StringCellValue == "Client")
                    {
                        continue;
                    }

                    if (dt.Columns.Count == 0 && row.GetCell(0).StringCellValue == "Client")
                    {
                        IRow headerRow = sheet.GetRow(rowIndex);
                        foreach (ICell headerCell in headerRow)
                        {
                            if (headerCell.ToString() != "")
                            {
                                dt.Columns.Add(headerCell.ToString());
                            }
                        }
                        rowIndex++;
                        continue;
                    }
                }


                DataRow dataRow = dt.NewRow();

                for (int cn = 0; cn < row.LastCellNum; cn++)
                {
                    ICell cell = row.GetCell(cn);
                    if (cell == null)
                    {                        
                        row.CreateCell(cn).SetCellValue("");
                    }
                }
               
                dataRow.ItemArray = row.Cells.Select(c => getCellValue(c)).ToArray();

                dt.Rows.Add(dataRow);

                rowIndex++;              
            }
            return dt;
        }

        public static string getCellValue(ICell cell)
        {
            string cellValue = "";
            
            switch (cell.CellType)
            {
                case CellType.Formula:
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try
                        {
                            cellValue = cell.DateCellValue.ToString("dd-MMM-yyyy");
                        }
                        catch (NullReferenceException e)
                        {
                            cellValue = DateTime.FromOADate(cell.NumericCellValue).ToString("dd-MMM-yyyy");
                        }                        
                    }
                    else
                    {
                        cellValue = cell.NumericCellValue.ToString();                       
                    }
                    break;
                case CellType.Blank:
                case CellType.Unknown:
                    cellValue = "";
                    break;
                default:
                    cellValue = cell.StringCellValue;                    
                    break;
            }
            
            return cellValue;
        }
    }
}
