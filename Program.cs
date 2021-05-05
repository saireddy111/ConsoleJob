using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace ConsoleJob
{
    class Program
    {
        static void Main(string[] args)
        {
            
            DataTable dt = GetDataTableFromExcel("D:\\Monthly File Format - Sample 4-15-2021.xlsx");

            List<MerchantModel> mList = new List<MerchantModel>();
            mList = (from DataRow dr in dt.Rows
                     select new MerchantModel()
                     {                        
                           ISOName = dr["ISO Name"].ToString(),
                           AgentSalesOfficeName = dr["Agent Sales Office Name"].ToString(),
                           AgentSalesOffice = dr["Agent Sales Office #"].ToString(),
                           MerchantMID = dr["Merchant # (MID)"].ToString(),
                           MerchantName = dr["Merchant Name"].ToString(),
                           DateBoardedOpen = dr["Date Boarded/Open"].ToString(),
                           DateClosed = dr["Date Closed"].ToString(),
                           ProgramType = dr["Program Type"].ToString(),
                           Platform = dr["Platform"].ToString(),
                           TransCount = dr["Trans Count"].ToString(),
                           SalesVolume = dr["Sales Volume"].ToString(),
                           SalesChannel = dr["Sales Channel"].ToString(),
                           Status = dr["Status"].ToString()

                     }).ToList();

            string xmlData = poco2Xml(mList);
            //string xml = ConvertDatatableToXML(dt);

        }

        private static string poco2Xml(object obj)
        {
            XmlSerializer serializer = new XmlSerializer(obj.GetType());
            StringBuilder result = new StringBuilder();
            using (var writer = XmlWriter.Create(result))
            {
                serializer.Serialize(writer, obj);
            }
            return result.ToString();
        }

        public static string ConvertDatatableToXML(DataTable dt)
        {
            MemoryStream str = new MemoryStream();
            dt.WriteXml(str, true);
            str.Seek(0, SeekOrigin.Begin);
            StreamReader sr = new StreamReader(str);
            string xmlstr;
            xmlstr = sr.ReadToEnd();
            return (xmlstr);
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

            IRow headerRow = sheet.GetRow(2);
            foreach (ICell headerCell in headerRow)
            {
                if (headerCell.ToString() != "")
                {
                    dt.Columns.Add(headerCell.ToString());
                }
            }

            int rowIndex = 0;
            foreach (IRow row in sheet)
            {
                if (rowIndex++ < 2) continue;
                DataRow dataRow = dt.NewRow();
                dataRow.ItemArray = row.Cells.Select(c =>
                c.CellType == CellType.Formula || c.CellType == CellType.Numeric ?
                (DateUtil.IsCellDateFormatted(c) ? getCellValue(c) : c.NumericCellValue.ToString()) :
                c.ToString()).ToArray();
                dt.Rows.Add(dataRow);
            }
            return dt;
        }

        public static string getCellValue(ICell cell)
        {
            string cellValue = "";
            try
            {
                cellValue = cell.DateCellValue.ToString("dd-MMM-yyyy");
            }
            catch (NullReferenceException e)
            {
                cellValue = DateTime.FromOADate(cell.NumericCellValue).ToString("dd-MMM-yyyy");
            }
            return cellValue;
        }
    }
}
