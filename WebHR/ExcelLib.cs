using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace WebHR
{
    class ExcelLib
    {
            public static DataTable ExcelToDataTable(string FileName1)
            {
                FileStream stream = File.Open(FileName1, FileMode.Open, FileAccess.Read);
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                // DataSet result = excelReader.AsDataSet();
                DataTableCollection table = result.Tables;
                DataTable resultTable = table[0];
                return resultTable;
            }

            public static List<Datacollection> dataCol = new List<Datacollection>();

            public static DataTable PopulateInCollection(string FileName)
            {
                DataTable table = ExcelToDataTable(FileName);

                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        Datacollection dtTable = new Datacollection()
                        {
                            rowNumber = row,
                            colName = table.Columns[col].ColumnName,
                            colValue = table.Rows[row - 1][col].ToString()
                        };
                        dataCol.Add(dtTable);
                    }
                }
                return table;
            }

            public static string ReadData(int rowNumber, string columnName)
            {
                try
                {

                    string data = (from colData in dataCol
                                   where colData.colName == columnName && colData.rowNumber == rowNumber
                                   select colData.colValue).SingleOrDefault();


                    return data.ToString();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception is : " + ex.Message);
                    return null;
                }
            }

        }

        public class Datacollection

        {
            public int rowNumber { get; set; }
            public string colName { get; set; }
            public string colValue { get; set; }
        }

}

