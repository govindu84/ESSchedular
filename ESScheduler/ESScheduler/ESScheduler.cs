using ClosedXML.Excel;
using ESScheduler.Models;
using Nest;
using ServiceReference;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace ESScheduler
{
    public class ESScheduler
    {
       
        private static ServiceReference.ConfigServiceClient wcfClient;

        public static List<ESAPP_Input> inputList;

        public ESScheduler()
        {
           wcfClient = new ConfigServiceClient((ConfigServiceClient.EndpointConfiguration.ConfigService_V3_BasicHttpService));
        }
        
      
        public static async void CombineInputData(string PathName, string URI)
        {
            try
            {
                 List<ESAPP_Input> revionicsDetails;
                 FileInfo fi = new FileInfo(PathName);
                if (fi.Extension == ".xls")
                {
                    File.Move(PathName, Path.ChangeExtension(PathName, ".xml"));                    
                    revionicsDetails = ExcelEngine.ImportExcelXML(Path.ChangeExtension(PathName, ".xml"));

                }
                else
                {
                    revionicsDetails = GetDataFromExcel(PathName);

                }

               
                //foreach (var item in revionicsDetails)
                //{
                //   // var priceModel = await GetPriceFromConfigService(item.OrderCode); //This is where we call another service
                //    item.RetailPrice = priceModel?.RetailPrice;
                //    item.TotalDiscount = priceModel?.TotalDiscount;
                //    item.SalePrice = priceModel?.SalePrice;
                //}

                var uri = new Uri(URI);
                var settings = new ConnectionSettings(uri);
                ElasticClient client = new ElasticClient(settings);
                settings.DefaultIndex("revionics");               
                client.IndexMany<ESAPP_Input>(revionicsDetails, null);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        //Get Details from WCF service
        private static async Task<SystemPriceModel> GetPriceFromConfigService(string orderCode)
        {
            SystemPriceModel priceModel;
            ProcessRequestRequest PRR = new ProcessRequestRequest();
            PRR.request = new ConfigRequest();
            PRR.request.ShipToCountry = LightweightProfileData.Country;
            PRR.request.CountryId = LightweightProfileData.Country;
            PRR.request.LanguageId = LightweightProfileData.Language;
            PRR.request.OrderCode = orderCode;
            PRR.request.CustomerSetId = LightweightProfileData.CustomerSet;

            ProcessRequestResponse response = await wcfClient.ProcessRequestAsync(PRR);
            priceModel = response?.ProcessRequestResult?.ConfigPrice;


            if (priceModel != null)
            {
                var systemPriceModel = new SystemPriceModel
                {
                    RetailPrice = (double)priceModel.RetailPrice,
                    SalePrice = (double)priceModel.SalePrice,
                    TotalDiscount = (double)priceModel.TotalDiscount

                };

                return systemPriceModel;
            }
            return null;
        }



        public static List<ESAPP_Input> GetDataFromExcel(string Pathname)
        {
            var xmlFile = Pathname; // Path.Combine(Environment.CurrentDirectory, "Data\\Input1.xlsx");
            using (var workBook = new XLWorkbook(xmlFile))
            {
                var workSheet = workBook.Worksheet(1);
                var firstRowUsed = workSheet.FirstRowUsed();
                var firstPossibleAddress = workSheet.Row(firstRowUsed.RowNumber()).FirstCell().Address;
                var lastPossibleAddress = workSheet.LastCellUsed().Address;

                // Get a range with the remainder of the worksheet data (the range used)
                var range = workSheet.Range(firstPossibleAddress, lastPossibleAddress).AsRange(); //.RangeUsed();
                                                                                                  // Treat the range as a table (to be able to use the column names)
                var table = range.AsTable();

                //Specify what are all the Columns you need to get from Excel
                var dataList = new List<string[]>
                 {
                     table.DataRange.Rows()
                         .Select(tableRow =>
                             tableRow.Field("Order Code")
                                 .GetString())
                         .ToArray(),
                     table.DataRange.Rows()
                         .Select(tableRow => tableRow.Field("Suggested Price").GetString())
                         .ToArray()
                     //table.DataRange.Rows()
                     //.Select(tableRow => tableRow.Field("Price Lock Start Date").GetString())
                     //.ToArray()
                 };
                //Convert List to DataTable
                var dataTable = ConvertListToDataTable(dataList);
                //To get unique column values, to avoid duplication
                var uniqueCols = dataTable.DefaultView.ToTable(true, "OrderCode");

                //Remove Empty Rows or any specify rows as per your requirement
                for (var i = uniqueCols.Rows.Count - 1; i >= 0; i--)
                {
                    var dr = uniqueCols.Rows[i];
                    if (dr != null && ((string)dr["OrderCode"] == "None"))
                        dr.Delete();
                }
                Console.WriteLine("Total number of unique solution number in Excel : " + uniqueCols.Rows.Count);

                inputList = new List<ESAPP_Input>();
                inputList = (from DataRow dr in dataTable.Rows
                             select new ESAPP_Input()
                             {
                                 OrderCode = Convert.ToString(dr["OrderCode"]),
                                 RevPrice = Convert.ToDouble(dr["RevPrice"]),
                                // Timestamp = !string.IsNullOrEmpty(dr["Timestamp"].ToString()) ? Convert.ToDateTime(dr["Timestamp"]) : DateTime.MinValue
                             }).ToList();

                return inputList;
            }
        }


        private static DataTable ConvertListToDataTable(IReadOnlyList<string[]> list)
        {
            var table = new DataTable("CustomTable");
            var rows = list.Select(array => array.Length).Concat(new[] { 0 }).Max();

            table.Columns.Add("OrderCode");
            table.Columns.Add("RevPrice");
           // table.Columns.Add("Timestamp");

            for (var j = 0; j < rows; j++)
            {
                var row = table.NewRow();
                row["OrderCode"] = list[0][j];
                row["RevPrice"] = list[1][j];
               // row["Timestamp"] = list[2][j];
                table.Rows.Add(row);
            }
            return table;
        }


    }

    public class ExcelEngine
    {
        public static List<ESAPP_Input> ImportExcelXML(string Path, bool hasHeaders=true, bool autoDetectColumnType=true)
        {

            ExcelEngine engine = new ExcelEngine();
            FileStream inputFileStream = File.Open(Path, FileMode.Open);

            XmlDocument doc = new XmlDocument();
            doc.Load(new XmlTextReader(inputFileStream));
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);

            nsmgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
            nsmgr.AddNamespace("x", "urn:schemas-microsoft-com:office:excel");
            nsmgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");

            DataSet ds = new DataSet();

            foreach (XmlNode node in doc.DocumentElement.SelectNodes("//ss:Worksheet", nsmgr))
            {
                DataTable dt = new DataTable(node.Attributes["ss:Name"].Value);
                ds.Tables.Add(dt);
                XmlNodeList rows = node.SelectNodes("ss:Table/ss:Row", nsmgr);
                if (rows.Count > 0)
                {
                    List<ColumnType> columns = new List<ColumnType>();
                    int startIndex = 0;
                    if (hasHeaders)
                    {
                        foreach (XmlNode data in rows[0].SelectNodes("ss:Cell/ss:Data", nsmgr))
                        {
                            columns.Add(new ColumnType(typeof(string)));//default to text
                            dt.Columns.Add(data.InnerText, typeof(string));
                        }
                        startIndex++;
                    }
                    if (autoDetectColumnType && rows.Count > 0)
                    {
                        XmlNodeList cells = rows[startIndex].SelectNodes("ss:Cell", nsmgr);
                        int actualCellIndex = 0;
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                        {
                            XmlNode cell = cells[cellIndex];
                            if (cell.Attributes["ss:Index"] != null)
                                actualCellIndex = int.Parse(cell.Attributes["ss:Index"].Value) - 1;

                            ColumnType autoDetectType = getType(cell.SelectSingleNode("ss:Data", nsmgr));

                            if (actualCellIndex >= dt.Columns.Count)
                            {
                                dt.Columns.Add("Column" + actualCellIndex.ToString(), autoDetectType.type);
                                columns.Add(autoDetectType);
                            }
                            else
                            {
                                dt.Columns[actualCellIndex].DataType = autoDetectType.type;
                                columns[actualCellIndex] = autoDetectType;
                            }

                            actualCellIndex++;
                        }
                    }
                    for (int i = startIndex; i < rows.Count; i++)
                    {
                        DataRow row = dt.NewRow();
                        XmlNodeList cells = rows[i].SelectNodes("ss:Cell", nsmgr);
                        int actualCellIndex = 0;
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                        {
                            XmlNode cell = cells[cellIndex];
                            if (cell.Attributes["ss:Index"] != null)
                                actualCellIndex = int.Parse(cell.Attributes["ss:Index"].Value) - 1;

                            XmlNode data = cell.SelectSingleNode("ss:Data", nsmgr);

                            if (actualCellIndex >= dt.Columns.Count)
                            {
                                for (int k = dt.Columns.Count; k < actualCellIndex; k++)
                                {
                                    dt.Columns.Add("Column" + actualCellIndex.ToString(), typeof(string));
                                    columns.Add(getDefaultType());
                                }
                                ColumnType autoDetectType = getType(cell.SelectSingleNode("ss:Data", nsmgr));
                                dt.Columns.Add("Column" + actualCellIndex.ToString(), typeof(string));
                                columns.Add(autoDetectType);
                            }
                            if (data != null)
                                row[actualCellIndex] = data.InnerText;

                            actualCellIndex++;
                        }

                        dt.Rows.Add(row);
                    }
                }
            }
            //return ds;

            List<ESAPP_Input> inputList = new List<ESAPP_Input>();
            inputList = (from DataRow dr in ds.Tables[0].Rows
                         select new ESAPP_Input()
                         {
                             OrderCode = Convert.ToString(dr["Order Code"]),
                             RevPrice = Convert.ToDouble(dr["Suggested Price"]),
                             // Timestamp = !string.IsNullOrEmpty(dr["Timestamp"].ToString()) ? Convert.ToDateTime(dr["Timestamp"]) : DateTime.MinValue
                         }).ToList();

            return inputList;
        }
        private static ColumnType getDefaultType()
        {
            return new ColumnType(typeof(String));
        }

        private static ColumnType getType(XmlNode data)
        {
            string type = null;
            if (data.Attributes["ss:Type"] == null || data.Attributes["ss:Type"].Value == null)
                type = "";
            else
                type = data.Attributes["ss:Type"].Value;

            switch (type)
            {
                case "DateTime":
                    return new ColumnType(typeof(DateTime));
                case "Boolean":
                    return new ColumnType(typeof(Boolean));
                case "Number":
                    return new ColumnType(typeof(Decimal));
                case "":
                    decimal test2;
                    if (data == null || String.IsNullOrEmpty(data.InnerText) || decimal.TryParse(data.InnerText, out test2))
                    {
                        return new ColumnType(typeof(Decimal));
                    }
                    else
                    {
                        return new ColumnType(typeof(String));
                    }
                default://"String"
                    return new ColumnType(typeof(String));
            }
        }
        struct ColumnType
        {
            public Type type;
            private string name;
            public ColumnType(Type type) { this.type = type; this.name = type.ToString().ToLower(); }
            public object ParseString(string input)
            {
                if (String.IsNullOrEmpty(input))
                    return DBNull.Value;
                switch (type.ToString())
                {
                    case "system.datetime":
                        return DateTime.Parse(input);
                    case "system.decimal":
                        return decimal.Parse(input);
                    case "system.boolean":
                        return bool.Parse(input);
                    default:
                        return input;
                }
            }
        }
    }

}
