using AutoMapper;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using GeographicDynamic_DAL.Models;
using GeographicDynamicWebAPI.Wrappers;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.EntityFrameworkCore.Query.Internal;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Azure.Core.HttpHeader;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using System.Data.SqlClient;


namespace GeographicDynamic_DAL.Repository
{
    public class ColumnNameRepository : IColumnName
    {

        private readonly IMapper _mapper;

        public ColumnNameRepository(IMapper mapper)
        {
            _mapper = mapper;
        }
        public Result<ColumnNameDTO> ColumnNameTransfer(string ExcelPath)

        {

            #region  gio 

            GeographicDynamicDbContext conn = new GeographicDynamicDbContext();
            
            List<string> columnNames = new List<string>();

            using (var workbook = new XLWorkbook(ExcelPath))
            {
                var worksheet = workbook.Worksheet(1);
                var firstRow = worksheet.FirstRow();

                // Count the number of columns in the first row
                int columnCount = 50;
                int col;
                // Loop through each column in the first row and add the column name to the list
                for (col = 1; col <= columnCount; col++)
                {
                    var cell = worksheet.Cell(1, col);
                    string columnName = cell.GetString();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        columnNames.Add(columnName);
                    }
                }
            }

           // List<ColumnNameDTO> ColumnNameDTOs = columnNames.Select(x => new ColumnNameDTO { ExcelName = x }).ToList();
             List<ColumnNameDTO> ColumnNameDTOs = columnNames.Select((x, index) => new ColumnNameDTO {Id = index + 1, ExcelName = x, ColN = index + 1 }).ToList();

            //// აქ მინდა დაიწეროს აფდეითი ექსელის წაკითხვის მერე და შეიყაროს ბაზაში 

            //foreach (var item in ColumnNameDTOs)
            //{
            //    var itemToUpdate = conn.ColumnNames.FirstOrDefault(x => x.ColN == item.ColN);
            //    if (itemToUpdate != null)
            //    {
            //        // If the item exists in the database, update its ColN value
            //        itemToUpdate.ColN = item.ColN;
            //    }
            //}

            //conn.SaveChanges();


            return new Result<ColumnNameDTO>
            {
                Success = true,
                Data = ColumnNameDTOs, // Assuming your DTO has a property named 'ColumnNames'
                StatusCode = System.Net.HttpStatusCode.OK
            };
            #endregion
            //// ითვლება პირველ როუში რამდენი ჩანაწერია 
            //int columnCount = firstRow.Columns.Count;



            //foreach (var col in firstRow.Columns)
            //{
            //    var columnNames = Convert.ToString((firstRow.Cells[1, col] as Excel.Range).Value);
            //    columnNames.Add(columnNames);
            //}



            //return new Result<ColumnNameDTO>
            //{
            //    Success = true,

            //    StatusCode = System.Net.HttpStatusCode.OK
            //};

        }

        public Result<ColumnNameDTO> ColumnNameTransferFromAccess(string AccessPath, string AccessSheetName)
        {
            var GeographicDynamicDbContext = new GeographicDynamicDbContext();
            List<string> columnNames = new List<string>();
            try
            {
                #region oledbConnection
                string connectionString = "";
                if (Path.GetExtension(AccessPath).ToLower().Trim() == ".mdb" && Environment.Is64BitOperatingSystem == false)
                {
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AccessPath;
                    connectionString = "Provider=Microsoft.Jet.OLEDBMicrosoft.Jet.OLEDB.4.0;Data Source=" + AccessPath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                }
                else
                {
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AccessPath;
                }

                using (var con = new OleDbConnection(connectionString))
                {
                    con.Open();
                    using (var cmd = new OleDbCommand("select * from " + AccessSheetName, con))
                    using (var reader = cmd.ExecuteReader(CommandBehavior.SchemaOnly))
                    {
                        var table = reader.GetSchemaTable();
                        var nameCol = table.Columns["ColumnName"];
                        foreach (DataRow row in table.Rows)
                        {
                            var test = row[nameCol];
                            columnNames.Add(row[nameCol].ToString());
                        }
                    }
                }

                #endregion

                List<ColumnNameDTO> ColumnNameDTOs = columnNames.Select(x => new ColumnNameDTO { AccessName = x }).ToList();


                return new Result<ColumnNameDTO>
                {
                    Success = true,
                    Data = ColumnNameDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<ColumnNameDTO>
                {
                    Success = false,
                    Data = null,
                    StatusCode = System.Net.HttpStatusCode.InternalServerError
                };
            }
        }
        public Result<ColumnNameDTO> GetSQLColumnNamesList()
        {
            try
            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();


                List<ColumnNameDTO> ColumnNameDTOs = geographicDynamicDbContext.ColumnNames.Where(x => x.Sqlname != null).Select(x => new ColumnNameDTO { Id = x.Id, Sqlname = x.Sqlname, ExcelName = x.ExcelName, DataType = x.DataType, AccessName = x.AccessName, IsAccessToExcel = x.IsAccessToExcel, GroupMethod = x.GroupMethod , ColN = x.ColN}).ToList();
                
                
                return new Result<ColumnNameDTO>
                {
                    Success = true,
                    Data = ColumnNameDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<ColumnNameDTO>
                {
                    Success = false,
                    Data = null,
                    StatusCode = System.Net.HttpStatusCode.InternalServerError
                };
            }
        }

        public Result<bool> SaveColumnName(List<ColumnNameDTO> columnNameDTO)
        {


            GeographicDynamicDbContext conn = new GeographicDynamicDbContext();
            try
            {
                var fieldsToUpdate = _mapper.Map<List<ColumnNameDTO>, List<ColumnName>>(columnNameDTO);

                foreach (var item in columnNameDTO)
                {

                    var existingFields = conn.ColumnNames.FirstOrDefault(x => x.Id == item.Id);
                    if (existingFields != null)
                    {
                        _mapper.Map(item, existingFields);
                        conn.SaveChanges();
                    }
                    else
                    {
                    }

                }
                List<ColumnName> columnNamesToInsert = fieldsToUpdate.Where(x => !conn.ColumnNames.Any(y => y.Id == x.Id)).ToList();


                string constr = "Data Source=WIN-IK4QOCMD77O;Initial Catalog=Geographic_Dynamic_DB;User Id=sa;Password=123;Trusted_Connection=True;TrustServerCertificate=True;";
                SqlCommand cmd = new SqlCommand();

                try
                {
                    foreach (var item in columnNamesToInsert)
                    {

                        using (SqlConnection connection = new SqlConnection(constr))
                        {
                            connection.Open();


                            string sqlCommandText = "";
                            switch (item.DataType)
                            {
                                case "ტექსტური":
                                    sqlCommandText = $"ALTER TABLE qarsafari ADD {item.Sqlname} NVARCHAR(255); " +
                                                     $"ALTER TABLE qarsafariGrouped ADD {item.Sqlname} NVARCHAR(255); " +
                                                     $"ALTER TABLE WindbreakMDB ADD {item.Sqlname} NVARCHAR(255);";
                                    break;
                                case "რიცხვითი":
                                    sqlCommandText = $"ALTER TABLE qarsafari ADD {item.Sqlname} FLOAT; " +
                                                     $"ALTER TABLE qarsafariGrouped ADD {item.Sqlname} FLOAT; " +
                                                     $"ALTER TABLE WindbreakMDB ADD {item.Sqlname} FLOAT;";

                                    break;
                                case "თარიღი":
                                    sqlCommandText = $"ALTER TABLE qarsafari ADD {item.Sqlname} DATE; " +
                                                     $"ALTER TABLE qarsafariGrouped ADD {item.Sqlname} DATE; " +
                                                     $"ALTER TABLE WindbreakMDB ADD {item.Sqlname} DATE;";
                                    break;
                            }
                            if (!string.IsNullOrEmpty(sqlCommandText))
                            {
                                using (SqlCommand sqlCommand = new SqlCommand(sqlCommandText, connection))
                                {
                                    sqlCommand.ExecuteNonQuery();
                                }
                            }


                        }
                        //ესენი Foreach-ში უნდა იყოს
                        conn.ColumnNames.Add(item);
                        conn.SaveChanges();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occurred: " + ex.Message);
                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით გადაინომრა"
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "ჩტო ტა ნიტო"
                };
            }



            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }

        public Result<int> DeleteRow(ColumnNameDTO Id)
        {
            try
            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();


                if (Id.Id != null && Id.Id != 0)
                {
                    var sqlName = geographicDynamicDbContext.ColumnNames.FirstOrDefault(x => x.Id == Id.Id);

                    if (sqlName != null)
                    {

                        string constr = "Data Source=WIN-IK4QOCMD77O;Initial Catalog=Geographic_Dynamic_DB;User Id=sa;Password=123;Trusted_Connection=True;TrustServerCertificate=True;";
                        SqlCommand cmd = new SqlCommand();
                        try
                        {
                            using (SqlConnection connection = new SqlConnection(constr))
                            {
                                connection.Open();


                                string sqlCommandText = $"DELETE FROM ColumnName WHERE ID = {Id.Id}";

                                if (!string.IsNullOrEmpty(sqlCommandText))
                                {
                                    using (SqlCommand sqlCommand = new SqlCommand(sqlCommandText, connection))
                                    {
                                        sqlCommand.ExecuteNonQuery();
                                    }
                                }


                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("An error occurred: " + ex.Message);
                        }
                    }
                }

                return new Result<int>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<int>
                {
                    Success = false,
                    Data = null,
                    StatusCode = System.Net.HttpStatusCode.InternalServerError
                };
            }
        }
    }
}
