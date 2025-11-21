using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamicWebAPI.Wrappers;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using ClosedXML.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace GeographicDynamic_DAL.Models.WindbreakFirstStepMethods
{
    public class WindbreakFirstStepMethods
    {
        #region Methods
        public Result<bool> FillIsUniqLitterNull()
        {
            try
            {
                GeographicDynamicDbContext GeographicDynamicDbContext = new GeographicDynamicDbContext();

                List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.ToList();

                foreach (var item in qarsafaris)
                {
                    if ((item.UniqId == null || item.LiterId == null) || (item.UniqId == 0 || item.LiterId == 0))
                    {
                        item.IsUniqLiterNull = "false";

                    }
                    if (item.UniqId != null && item.LiterId != null)
                    {
                        item.IsUniqLiterNull = "true";
                    }



                    GeographicDynamicDbContext.SaveChanges();

                }


                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა IsUniqLitterNull ველის შევსების დროს" + ex.Message
                };
            }

        }

        ///////აქ იკითხება ძირი ექსელი და შედის sql ბაზაში 
        public Result<bool> ExcelisWakitxva(ExcelReadDTO excelReadDTO)
        {
            GeographicDynamicDbContext GeographicDynamicDbContext = new GeographicDynamicDbContext();
            var test = excelReadDTO.UnicIDStartNumber;
            var test1 = excelReadDTO.ExcelDestinationPath;
            var ExcelPath = excelReadDTO.ExcelPath;
            var test3 = excelReadDTO.AccessFilePath;
            var municipality = excelReadDTO.ProjectNameID;

            try
            {
                using (var workbook = new XLWorkbook(ExcelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var lastRow = worksheet.LastRowUsed();
                    int colCount = lastRow != null ? lastRow.RowNumber() : 0;

                    //ExcelProgressBar.Minimum = 1;
                    //ExcelProgressBar.Maximum = colCount - 1;
                    //ExcelProgressBar.Step = 1;

                    //if (colCount == ExcelProgressBar.Maximum)
                    //{
                    //    ExcelProgressBar.Maximum = 100;
                    //}

                    // წინასწარ ცხრილის გასუფთავება მანამ ჩანაწერებს შევიტანთ

                    // ----აქეეედააააან
                    GeographicDynamicDbContext.Qarsafaris.ExecuteDelete();
                    Type myType = typeof(Qarsafari);
                    //იტერაცია ექსელის ფაილში
                    for (int i = 2; i <= colCount; i++) // colCount tu sworad wakikitxavs
                    {
                        Qarsafari qarsafari = new Qarsafari();
                        var row = worksheet.Row(i);

                        // თუ უნიკიდ ან ლიტერ აიდი ცარიელია მაშინ ჩაიწერება false თუ არაა ცარიელი მაშინ true
                        var cellACell = row.Cell("A");
                        var cellBCell = row.Cell("B");
                        string cellA = cellACell.IsEmpty() ? string.Empty : (cellACell.DataType == XLDataType.Number ? cellACell.GetValue<double>().ToString() : cellACell.GetValue<string>());
                        string cellB = cellBCell.IsEmpty() ? string.Empty : (cellBCell.DataType == XLDataType.Number ? cellBCell.GetValue<double>().ToString() : cellBCell.GetValue<string>());
                        if (String.IsNullOrEmpty(cellA) || String.IsNullOrEmpty(cellB))
                        {
                            qarsafari.IsUniqLiterNull = "false";
                        }
                        else
                        {
                            qarsafari.IsUniqLiterNull = "true";
                        }

                        foreach (var columnName in GeographicDynamicDbContext.ColumnNames)
                        {
                            /////////ამ იფ სთეითმენთით ვახტებით WoodyPlantQuantity სვეტს რადგან არ წავიკითხოთ შემდეგ მეთოდში რომ შეივსოს და არ გადაიწეროს 

                            //if (columnName.Sqlname == "WoodyPlantQuantity")
                            //{
                            //    continue;
                            //}
                            //Get cell type
                            if (columnName.ColN != null)
                            {
                                var cell = row.Cell(columnName.ColN.Value);
                                if (!cell.IsEmpty())
                                {
                                    PropertyInfo propertyInfo = typeof(Qarsafari).GetProperty(columnName.Sqlname);
                                    if (propertyInfo != null)
                                    {
                                        object cellValue = null;
                                        
                                        // Get value based on cell data type
                                        if (cell.DataType == XLDataType.Number)
                                        {
                                            cellValue = cell.GetValue<double>();
                                        }
                                        else if (cell.DataType == XLDataType.Text)
                                        {
                                            cellValue = cell.GetValue<string>();
                                        }
                                        else if (cell.DataType == XLDataType.DateTime)
                                        {
                                            cellValue = cell.GetValue<DateTime>();
                                        }
                                        else if (cell.DataType == XLDataType.Boolean)
                                        {
                                            cellValue = cell.GetValue<bool>();
                                        }
                                        else
                                        {
                                            // Try to get as string for other types
                                            try
                                            {
                                                cellValue = cell.GetValue<string>();
                                            }
                                            catch
                                            {
                                                // If string conversion fails, try double
                                                try
                                                {
                                                    cellValue = cell.GetValue<double>();
                                                }
                                                catch
                                                {
                                                    // Skip this cell if we can't convert it
                                                    continue;
                                                }
                                            }
                                        }

                                        if (cellValue != null)
                                        {
                                            Type cellType = cellValue.GetType();
                                            Type propertyType = propertyInfo.PropertyType;
                                            
                                            // Handle nullable types
                                            if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                                            {
                                                propertyType = Nullable.GetUnderlyingType(propertyType);
                                            }

                                            // Handle conversion based on property type
                                            if (propertyType == typeof(double) || propertyType == typeof(double?))
                                            {
                                                if (cellValue is double doubleVal)
                                                {
                                                    propertyInfo.SetValue(qarsafari, doubleVal);
                                                }
                                                else if (double.TryParse(cellValue.ToString(), out double parsedDouble))
                                                {
                                                    propertyInfo.SetValue(qarsafari, parsedDouble);
                                                }
                                            }
                                            else if (propertyType == typeof(string))
                                            {
                                                propertyInfo.SetValue(qarsafari, cellValue.ToString());
                                            }
                                            else if (propertyType == typeof(DateTime) || propertyType == typeof(DateTime?))
                                            {
                                                if (cellValue is DateTime dateTimeVal)
                                                {
                                                    propertyInfo.SetValue(qarsafari, dateTimeVal);
                                                }
                                                else if (DateTime.TryParse(cellValue.ToString(), out DateTime parsedDateTime))
                                                {
                                                    propertyInfo.SetValue(qarsafari, parsedDateTime);
                                                }
                                            }
                                            else if (propertyType == typeof(int) || propertyType == typeof(int?))
                                            {
                                                if (cellValue is int intVal)
                                                {
                                                    propertyInfo.SetValue(qarsafari, intVal);
                                                }
                                                else if (int.TryParse(cellValue.ToString(), out int parsedInt))
                                                {
                                                    propertyInfo.SetValue(qarsafari, parsedInt);
                                                }
                                                else if (cellValue is double doubleVal && doubleVal == Math.Truncate(doubleVal))
                                                {
                                                    propertyInfo.SetValue(qarsafari, (int)doubleVal);
                                                }
                                            }
                                            else if (propertyType == typeof(bool) || propertyType == typeof(bool?))
                                            {
                                                if (cellValue is bool boolVal)
                                                {
                                                    propertyInfo.SetValue(qarsafari, boolVal);
                                                }
                                                else if (bool.TryParse(cellValue.ToString(), out bool parsedBool))
                                                {
                                                    propertyInfo.SetValue(qarsafari, parsedBool);
                                                }
                                            }
                                            else
                                            {
                                                // Try direct assignment if types match
                                                if (propertyType.IsAssignableFrom(cellType))
                                                {
                                                    propertyInfo.SetValue(qarsafari, cellValue);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        GeographicDynamicDbContext.Qarsafaris.Add(qarsafari); // ახალი ობიექტის დამატება ბაზაში
                        GeographicDynamicDbContext.SaveChanges(); // ცვლილებების შენახვა

                        //label2.Text = Convert.ToString($"{i - 1} / {colCount - 1}");
                        //GeographicDynamicDbContext.PerformStep();
                    }
                }

                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                return new Result<bool> { Success = false, StatusCode = System.Net.HttpStatusCode.BadGateway, Message = "შეცდომა მოხდა " + ex.Message };
            }
        }

        ///////////////////// მეთოდი კრიბავს ხეხილს კარგმდგომარეობაში გაჩეხილი და გამხმარი და შედეგი იწერება მცენარეების რაოდენობის ველში 
        public Result<bool> fillAmountOfSpeces(ExcelReadDTO excelReadDTO)
        {

            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
            try
            {


                foreach (var item in geographicDynamicDbContext.Qarsafaris.ToList())
                {

                    var InGoodCondition = (item.InGoodCondition != null ? item.InGoodCondition : 0);
                    var Rampike = (item.Rampike != null ? item.Rampike : 0);
                    var ChoppedDown = (item.ChoppedDown != null ? item.ChoppedDown : 0);
                    item.WoodyPlantQuantity = InGoodCondition + Rampike + ChoppedDown;
                    ///////////აქ ჩაემატა ასევე etapiId და ProjectId ველების შევსება რადგან არქივში გადატანისას ამ ველების მიხედვით ვახდენთ ცვლილებებს
                    item.EtapiId = excelReadDTO.EtapiID;
                    item.ProjectId = excelReadDTO.ProjectID;
                }
                geographicDynamicDbContext.SaveChanges();
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა ხეხილის რაოდენობების ჩაწერა "
                };

            }
            catch
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "წარუმატებლად დასრულდა ხეხილის რაოდენობების ჩაწერა  "

                };
            }
        }


        ////აქ უნდა შემოწმდეს ლიტერი უნიკიდი  თუ მეორედება ექსელში მაშინ აღარ უდნა გააგრძელოს პროცესი 
        /// 

        public Result<double?> ShemowmebaUnicLiterExcelshi()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
            List<string> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.Where(x => x.IsUniqLiterNull == "true")
                                                                                        .OrderBy(m => m.UniqId)
                                                                                        .Select(q => $"{q.UniqId}-{q.LiterId}")
                                                                                        .ToList();
            try
            {
                return new Result<double?>
                {
                    Success = true,
                    // Data = uniqIdsNotInAccessList,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა შემოწმება Excel-ში UniqId-ის "
                };

            }
            catch
            {
                return new Result<double?>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "წარუმატებლად დასრულდა შემოწმება Excel-ში UniqId-ის "

                };
            }
        }
        //// ეს ფუქნცია ამოწმებს excel და access ცხრილებს და ადარებს UNIQID ებს თუ ემთხვევა ერთმანეთს 



        #region მოწმდება MDB Excel და fotoebi და გამოქავს შედეგი თუ სადმე ცხრილებს შორის დუბლიკატია ანდა რამე ზედმეტი ან ნაკლებია
        public Result<string?> ShemowmebaAccessExcelUnicLiterDublicats()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            string uniqIdsNotInAccessList = "";
            List<string> uniqIdsNotInAccessListActual = new List<string>();
            try
            {

                #region ALEKS

                //List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(m => m.IsUniqLiterNull == "true").ToList();
                //List<WindbreakMdb> windbreakMdbs = geographicDynamicDbContext.WindbreakMdbs.ToList();

                //bool emtxveva = true;
                //foreach (var excel in qarsafaris)
                //{
                //    foreach (var mdb in windbreakMdbs)
                //    {
                //        if (mdb.UniqId == excel.UniqId && mdb.LiterId == excel.LiterId)
                //        {
                //            emtxveva = false;
                //        }
                //    }
                //}


                #endregion


                #region gio
                //List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();
                //List<double?> AccessList = geographicDynamicDbContext.WindbreakMdbs.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();

                //foreach (var item in AccessList)
                //{
                //    if (!distinctUniqIds.Contains(item))
                //    {
                //        uniqIdsNotInAccessList.Add(item);
                //    }
                //}

                //if (uniqIdsNotInAccessList.Count > 0)
                //{
                //    return new Result<double?>
                //    {
                //        Success = false,
                //        Data = uniqIdsNotInAccessList,
                //        StatusCode = System.Net.HttpStatusCode.BadGateway,
                //        Message = "მოხდა შეცდომა ! Access და Excel UniqId-ები არ ემთხვევა ერთმანეთს !"
                //    };
                //}

                //return new Result<double?>
                //{
                //    Success = true,
                //    Data = uniqIdsNotInAccessList,
                //    StatusCode = System.Net.HttpStatusCode.OK,
                //    Message = "წარნატებით დასრულდა შემოწმება Access და Excel UniqId-ის "
                //};
                #endregion


                List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(m => m.IsUniqLiterNull == "true").Select(x => new Qarsafari { UniqId = x.UniqId, LiterId = x.LiterId }).ToList();

                var duplicates = qarsafaris.GroupBy(q => new { q.UniqId, q.LiterId }).Where(g => g.Count() > 1).SelectMany(g => g);

                if (duplicates.Any())
                {
                    //Console.WriteLine("Duplicates found:");
                    foreach (var duplicate in duplicates)
                    {
                        uniqIdsNotInAccessList += $"{duplicate.LiterId}-{duplicate.UniqId})";
                    }
                    return new Result<string?>
                    {
                        Success = false,
                        //Data = uniqIdsNotInAccessList,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = "მოხდა შეცდომა ! Excel UniqId  !: " + uniqIdsNotInAccessList
                    };
                }

                List<WindbreakMdb> windbreakMdbs = geographicDynamicDbContext.WindbreakMdbs.Select(x => new WindbreakMdb { UniqId = x.UniqId, LiterId = x.LiterId }).ToList();

                var duplicatesMDB = qarsafaris.GroupBy(q => new { q.UniqId, q.LiterId }).Where(g => g.Count() > 1).SelectMany(g => g);

                if (!duplicates.Any())
                {

                    //Console.WriteLine("Duplicates found:");
                    foreach (var duplicate in duplicatesMDB)
                    {
                        uniqIdsNotInAccessList += $"{duplicate.LiterId}-{duplicate.UniqId})";
                        return new Result<string?>
                        {
                            Success = false,
                            //Data = uniqIdsNotInAccessList,
                            StatusCode = System.Net.HttpStatusCode.BadGateway,
                            Message = "მოხდა შეცდომა ! Excel UniqId  !"
                        };
                    }



                    List<Qarsafari> resultList = qarsafaris.Where(u => windbreakMdbs.Any(l => l.LiterId == u.LiterId && l.UniqId == u.UniqId)).ToList();

                    /////////ესენი დაკომენტარებული იო და ახლა გასატესტია
                    foreach (var excel in qarsafaris)
                    {
                        bool existsInList = windbreakMdbs.Any(x => x.UniqId == excel.UniqId && x.LiterId == excel.LiterId);
                        if (!existsInList)
                        {
                            uniqIdsNotInAccessListActual.Add(string.Concat(excel.UniqId, "-", excel.LiterId, "excel"));
                            //uniqIdsNotInAccessListActual.Add(string.Concat(excel.UniqId.ToString(), "-", excel.LiterId.ToString(), "excel"));

                        }
                    }
                    foreach (var access in windbreakMdbs)
                    {
                        bool existsInList = qarsafaris.Any(x => x.UniqId == access.UniqId && x.LiterId == access.LiterId);
                        if (!existsInList)
                        {
                            uniqIdsNotInAccessListActual.Add(string.Concat(access.UniqId, "-", access.LiterId, "access"));
                        }
                    }
                    resultList = qarsafaris.Where(u => windbreakMdbs.Any(l => l.LiterId == u.LiterId && l.UniqId == u.UniqId)).ToList();


                    if (uniqIdsNotInAccessListActual.Count != 0)
                    {
                        string? concatenatedString = "";
                        foreach (var item in resultList)
                        {
                            concatenatedString += $"{item.LiterId}-{item.UniqId}";
                        }

                        return new Result<string?>
                        {
                            Success = false,
                            Data = uniqIdsNotInAccessListActual.ToList(),
                            StatusCode = System.Net.HttpStatusCode.BadGateway,
                            Message = "მოხდა შეცდომა ! Excel და Access რაოდენობა არ ემთხვევა!"
                        };
                    }







                    return new Result<string?>
                    {
                        Success = true,
                        //Data = uniqIdsNotInAccessList,
                        StatusCode = System.Net.HttpStatusCode.OK,
                        Message = "წარნატებით დასრულდა შემოწმება access და excel uniqid-ის "
                    };
                }

                return new Result<string?> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };

            }
            catch
            {
                return new Result<string?>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "წარუმატებლად შესრულდა შემოწმება access და excel uniqid-ის "
                };
            }
        }
        #endregion
        // ფუნქცია გამოიყენება რომ შეავსოს ველები სადაც გვიწერია პროექტის(მუნიციპალიტეტის) დასახელება და ეტაპის ნუმერაცია 
        public Result<string?> FillProjectEtapiIDS(int ProjectNameID, int EtapiID)
        {
            try
            {
                GeographicDynamicDbContext GeographicDynamicDbContext = new GeographicDynamicDbContext();

                List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.ToList();


                foreach (var item in qarsafaris)
                {
                    item.ProjectId = ProjectNameID;
                    item.EtapiId = EtapiID;
                    GeographicDynamicDbContext.SaveChanges();
                }


                return new Result<string?> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };

            }
            catch
            {
                return new Result<string?>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "წარუმატებლად შესრულდა ProjectID da EtapiID ჩაწერა "
                };
            }
        }



        //ეს ფუნქცია მიდის და ქარსაფარის ცხრილში სახეობების მიხედვით აკეთებს ვარჯის ფართების ჩაწერას
        public Result<bool> ChaweraVarjisParti(int ProjectNameID)
        {
            try
            {
                // Creating a GeographicDynamicDbContext instance
                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<VarjisFarti> varjisFartis = GeographicDynamicDbContext.VarjisFartis.Where(x => x.AreaNameId == ProjectNameID).ToList();
                foreach (var item in varjisFartis)
                {
                    var saxeobaName = GeographicDynamicDbContext.Dictionaries.FirstOrDefault(m => m.Id == item.SaxeobaId).Name;
                    List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.Where(x => x.WoodyPlantSpecies == saxeobaName).ToList();
                    foreach (var qarsafariItem in qarsafaris)
                    {
                        qarsafariItem.VarjisFarti = item.VarjisFarti1;

                    }
                    GeographicDynamicDbContext.SaveChanges();
                }

                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                // Returning failure result with error message
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ვარჯის ფართების გადათვლისას: " + ex.Message
                };
            }
        }


        //ვარჯის ფართების შემოწმება სადაც ხეხილი წერია და ვარჯის ფართი არა 
        public Result<bool> CheckerOfVarjisFartiandSaxeoba()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.ToList();
            try
            {
                foreach (var item in qarsafaris)
                {
                    if (item.VarjisFarti == null && item.WoodyPlantSpecies != null)
                    {
                        return new Result<bool>
                        {
                            Success = false,
                            StatusCode = System.Net.HttpStatusCode.BadRequest,
                            Message = "ვარჯისფართი არ ჩაიწერა სადაც სახეობა გვაქ! "
                        };
                    }

                }
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "შემოწმებისას მოხდა შეცდომა ვარჯისფართი არ წერია სადაც სახეობა გვაქ ! " + ex.Message
                };
            }
        }


        // ეშვება მეთოდი იმისთვის რომ LITER_ID და UNIQ-ID შეერთდეს და ჩაიწეროს UID-ში
        public Result<bool> UIDReplaceExcel()
        {
            try
            {
                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.ToList();
                foreach (var item in qarsafaris)
                {
                    item.Uid = item.LiterId.ToString() + item.UniqId.ToString();
                    //item.Uid = String.Concat(item.LiterId, item.UniqId);
                    GeographicDynamicDbContext.SaveChanges();
                }
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა UID replace-ის დროს" + ex.Message
                };
            }
        }



        // ფუნქცია გამოიყენება რომ წაიკითხოს Access ფაილი და შეყაროს SQL ბაზაში 
        public Result<bool> AccessWakitxva(string AccessFilePath, string AccessShitName)
        {
            var GeographicDynamicDbContext = new GeographicDynamicDbContext();
            var AccessPath = AccessFilePath;
            var AccessShitN = AccessShitName;

            GeographicDynamicDbContext.WindbreakMdbs.ExecuteDelete();

            #region  OleDbConnection for Access

            //OleDbConnection

            //string connectionString = @"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\Users\gioch\OneDrive\Desktop\GEOGraphics\test.accdb";
            //string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\gioch\\OneDrive\\Desktop\\GEOGraphics\Dedoplistskaro.mdb";
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

            #endregion

            string strSQL = "SELECT * FROM " + AccessShitN;
            // Create a connection    
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Create a command and set its connection    
                OleDbCommand command = new OleDbCommand(strSQL, connection);
                // Open the connection and execute the select command.    
                try
                {

                    // Open connecton    
                    connection.Open();
                    // Execute command    
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            WindbreakMdb windbreakMdb = new WindbreakMdb();// იყო ზევით და არ მუშაობდა რადგან ყოველ ჯერზე ახალი შექმნას და არ იყოს შევსებული 
                            foreach (var columnName in GeographicDynamicDbContext.ColumnNames)
                            {


                                //Get cell type
                                if (columnName.AccessName != null)
                                {
                                    object cellValue = reader[columnName.AccessName];
                                    if (cellValue != null)
                                    {
                                        Type cellType = cellValue.GetType();
                                        PropertyInfo propertyInfo = typeof(WindbreakMdb).GetProperty(columnName.Sqlname);
                                        if (propertyInfo != null)
                                        {

                                            // Handle conversion based on cell type
                                            if (cellType == typeof(System.Single))
                                            {
                                                Double? doubleValue = Convert.ToDouble(cellValue);
                                                propertyInfo.SetValue(windbreakMdb, doubleValue);
                                            }
                                            else if (cellType == typeof(System.Int32))
                                            {
                                                Double? intValue = Convert.ToDouble(cellValue);
                                                propertyInfo.SetValue(windbreakMdb, intValue);
                                            }
                                            else if (cellType == typeof(string))
                                            {
                                                propertyInfo.SetValue(windbreakMdb, cellValue);
                                            }
                                            else if (cellType == typeof(DateTime))
                                            {
                                                DateTime dateTimeValue;
                                                if (DateTime.TryParse((string)cellValue, out dateTimeValue))
                                                {
                                                    propertyInfo.SetValue(windbreakMdb, dateTimeValue);
                                                }
                                                // Handle DateTime conversion if necessary
                                            }
                                            // Add other type conversions as necessary
                                        }
                                    }

                                }

                            }

                            GeographicDynamicDbContext.WindbreakMdbs.Add(windbreakMdb);
                            GeographicDynamicDbContext.SaveChanges();
                            //// ველების წაკითხვის და ბაზაში გაშვების/დამახსოვრების ციკლი 
                            //while (reader.Read())
                            //{
                            //    WindbreakMdb windbreakMdbs = new WindbreakMdb();

                            //    windbreakMdbs.UniqId = (float?)reader["UNIQ_ID"];
                            //    windbreakMdbs.LiterId = (float?)reader["LITER_ID"];
                            //    windbreakMdbs.AdmMun = reader["Adm_Mun"].ToString();
                            //    windbreakMdbs.CityTownVillage = reader["City_Town_Village"].ToString();
                            //    windbreakMdbs.LandAreaSqM = (float?)reader["Land_Area_Sq_m"];
                            //    windbreakMdbs.LandAreaHa = (float?)reader["Land_Area_Ha"];
                            //    windbreakMdbs.LegalPerson = reader["Legal_person"].ToString();
                            //    windbreakMdbs.Note = reader["Note_"].ToString();
                            //    windbreakMdbs.Date = reader["Date_"].ToString();
                            //    windbreakMdbs.GisOperator = reader["Gis_Operator"].ToString();
                            //    windbreakMdbs.DaTe1 = reader["DaTe_1"].ToString();
                            //    windbreakMdbs.OverlapCadCode = reader["Overlap_CAD_COD"].ToString();
                            //    windbreakMdbs.Owner = reader["Owner"].ToString();



                            //    // ამატებს SQL ბაზაში და ამახსოვრებს ცვლილებებს 
                            //    GeographicDynamicDbContext.WindbreakMdbs.Add(windbreakMdbs);
                            //    GeographicDynamicDbContext.SaveChanges();

                            //    //Console.WriteLine("{0} {1}", reader["Name"].ToString(), reader["Address"].ToString());
                            //}


                        }
                    }
                }
                catch (Exception ex)
                {
                    return new Result<bool>
                    {
                        Success = false,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = "აქსესის წაკითხვის მოხდა შეცდომა ! " + ex.Message
                    };
                }
                // The connection is automatically closed becasuse of using block.    
            }

            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }



        //ამ ფუნქციაზი ხდება შემოწმება UniqID - ების ექსელში და აქსესში 
        // ფუნქცია გამოიყენება რომ მოხდეს Access ფაილიდან წაკითხული მონაცემები და გადასული ინფორმაციის UID ველის შევსება ლიტერის და უნიკაიდის კონკატენაციით 
        public Result<bool> UIDReplaceAccess()
        {
            try
            {

                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<WindbreakMdb> windbreakMdbs = GeographicDynamicDbContext.WindbreakMdbs.ToList();
                foreach (var item in windbreakMdbs)
                {
                    item.Uid = String.Concat(item.LiterId, item.UniqId);
                    GeographicDynamicDbContext.SaveChanges();

                }
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }

            catch (Exception ex)
            {

                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა UID replace-ის დროს" + ex.Message
                };
            }
        }



        // ფუნქცია გამოიყენება რომ დაგაიწეროს Access ფაილიდან საჭირო მონაცემები Excel-ში 
        public Result<bool> UpdateFromAccessToExcell()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            List<WindbreakMdb> AccessList = geographicDynamicDbContext.WindbreakMdbs.ToList();
            List<Qarsafari> ExcelList = geographicDynamicDbContext.Qarsafaris.ToList();
            try
            {
                foreach (var excel in ExcelList)
                {
                    if (excel.IsUniqLiterNull == "true")
                    {
                        WindbreakMdb access = geographicDynamicDbContext.WindbreakMdbs.FirstOrDefault(x => x.LiterId == excel.LiterId && x.UniqId == excel.UniqIdOld);
                        if (access != null)
                        {
                            foreach (var ColumnName in geographicDynamicDbContext.ColumnNames)
                            {

                                if (ColumnName.IsAccessToExcel == true)
                                {
                                    PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(ColumnName.Sqlname);

                                    if (propertyInfoQarsafari != null)
                                    {
                                        PropertyInfo propertyInfoWindbreakMDB = typeof(WindbreakMdb).GetProperty(ColumnName.Sqlname);
                                        object propertyValueWindbreakMDB = propertyInfoWindbreakMDB.GetValue(access);
                                        // თუ ცარიელია ვამოწმებთ და ვწერთ null -ს 
                                        if (propertyValueWindbreakMDB != null)
                                        {

                                            //აქ სადაც double ია მოაქვს System.Single ამიტომ ვამოწმებთ
                                            if (propertyValueWindbreakMDB.GetType() == typeof(System.Single))
                                            {
                                                Double? intValue = Convert.ToDouble(propertyValueWindbreakMDB);
                                                propertyInfoQarsafari.SetValue(excel, intValue, null);
                                            }
                                            else // სხვა შემთხვევაში არის string ან date ან bit
                                            {
                                                propertyInfoQarsafari.SetValue(excel, propertyInfoWindbreakMDB.GetValue(access, null), null);
                                            }
                                        }
                                        else
                                        {
                                            propertyInfoQarsafari.SetValue(excel, null, null);
                                        }
                                    }
                                }
                            }

                            geographicDynamicDbContext.SaveChanges();
                        }
                    }
                }
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა Access-დან Excel-ში "
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა Access-დან Excel-ში გადაწერის დროს" + ex.Message
                };
            }




        }

        // ითვლება პროცენტული მაჩვენებელი ხეხილის თუ რამდენია კარგ მდგომარეობაში და ასე შემდეგ 
        public Result<bool> QarsafariProcentisDatvla()

        {

            try

            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                List<Qarsafari> ExcelList = geographicDynamicDbContext.Qarsafaris.ToList();

                foreach (var Excel in ExcelList)

                {
                    Excel.ChoppedDownQuantity = Excel.ChoppedDown;
                    if (Excel.ChoppedDown != 0 && Excel.ChoppedDown != null)
                    {
                        Excel.ChoppedDown = (Excel.ChoppedDown / Excel.WoodyPlantQuantity) * 100;
                        Excel.Gachexili = (Excel.ChoppedDown / Excel.WoodyPlantQuantity) * 100;
                    }
                    if (Excel.InGoodCondition != 0 && Excel.InGoodCondition != null)
                    {
                        Excel.InGoodCondition = (Excel.InGoodCondition / Excel.WoodyPlantQuantity) * 100;
                    }
                    if (Excel.Rampike != 0 && Excel.Rampike != null)
                    {
                        Excel.Rampike = (Excel.Rampike / Excel.WoodyPlantQuantity) * 100;
                    }

                    geographicDynamicDbContext.SaveChanges();

                }

                return new Result<bool>

                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა ქარსაფარში პროცენტის დათვლა "
                };
            }

            catch (Exception ex)

            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ქარსაფარში პროცენტის დათვლის დროს" + ex.Message
                };
            }

        }
        //ფუნქცია გამოიყენება რომ დაგაინომროს UNIQ_ID ები ქარსაფარის ცხრილში 
        public Result<bool> QarsafariGadanomrva(int UnicIDStartNumber)
        {
            var geographicDynamicDbContext = new GeographicDynamicDbContext();
            try
            {
                #region ALEKS
                List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.ToList();
                //UniqId გადაგვაქვს UniqIdOld -ში ძველი უნიკიდის შესანახად
                foreach (var qarsafari in qarsafaris)
                {
                    if (qarsafari.IsUniqLiterNull == "true")
                    {
                        qarsafari.UniqIdOld = Convert.ToInt32(qarsafari.UniqId);
                    }
                }
                geographicDynamicDbContext.SaveChanges();

                //გადანომრვა
                var newUniqueID = UnicIDStartNumber - 1;
                //გლობალურად ვინათავთ ლიტერაიდის რომ შემდეგ იტერაციაში გამოვიყენოთ 
                Double? literid = null;
                foreach (var qarsafari in qarsafaris)
                {
                    if (qarsafari.IsUniqLiterNull == "true")
                    {
                        newUniqueID++;
                        //აქ იღებს ლიტერაიდი მნიშვნელობას როდესაც ზედა if პირობა სრულდება მაშინ იცვლის მნიშვნელობას 
                        literid = qarsafari.LiterId;
                    }
                    qarsafari.UniqId = newUniqueID;

                    //აქ უკვე იწერება ქარსაფარში 
                    qarsafari.LiterId = literid;
                }
                //ვიმახსოვრებთ შედეგებს 
                geographicDynamicDbContext.SaveChanges();
                #endregion
                #region GIO
                //List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.OrderBy(x => x.LiterId).ThenBy(x => x.UniqId).ToList();

                //foreach (var item in qarsafaris)
                //{
                //    var implement = UnicIDStartNumber;
                //    item.UniqIdOld = Convert.ToInt32(item.UniqId);
                //    if (item.Municipality != null)
                //    {
                //        item.UniqId = implement;
                //    }

                //    foreach (var item1 in qarsafaris)
                //    {
                //        if(item1.Municipality == null)
                //        {
                //            //qarsafaris.FirstOrDefault()
                //        }
                //    }
                //    implement++;
                //}
                #endregion
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა გადანომვრის პროცესი"
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა გადანომვრის პროცესის დროს" + ex.Message
                };
            }
        }
        // აქ გვჭირდება შემმოწმება ფუნქციის ჩაწერა რომელიც გადაამოწმებს თუ სადმე ხეხილი მეორდება უბანზე 
        public Result<bool> QarsafariXexilisShemowmeba()

        {

            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            try
            {
                List<Qarsafari> qarsafariList = geographicDynamicDbContext.Qarsafaris.ToList();
                List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();
                var count = 0;
                foreach (var uniqueID in distinctUniqIds)
                {
                    List<string> xexilebiList = new List<string>();
                    List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(x => x.UniqId == uniqueID).ToList();
                    foreach (var item in qarsafaris)
                    {
                        if (!string.IsNullOrEmpty(item.WoodyPlantSpecies))
                        {
                            xexilebiList.Add(item.WoodyPlantSpecies);
                        }
                    }
                    if (xexilebiList != null && xexilebiList.Count != xexilebiList.Distinct().Count())
                    {
                        count++;
                    }
                }
                if (count > 0)
                {
                    return new Result<bool>
                    {
                        Success = false,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = "მოხდა შეცდომა! ხეხილის ჯიში მეორდება"
                    };

                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა გადანომვრის პროცესი"
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა გადანომვრის პროცესის დროს" + ex.Message
                };
            }
        }
        //ფუნქცია გამოიყენება რომ დაიგრუპოს ხეხილის სახეობები და ამასთან მიყვეს სხვა პროცედურებიც რაც დაგრუპვაში შედის (პატარა ექსელი) 
        public Result<bool> QarsafariToQarsafariGrouped()
        {
            try
            {

                //თავიდან უნდა გავასუფთავოთ qarsafariGrouped ცხრილი

                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                geographicDynamicDbContext.QarsafariGroupeds.ExecuteDelete();
                // List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.Where(m => m.UniqId ).Select(q => q.UniqId).Distinct().ToList(); ერთი კონკრეტული როუს გასატესტად 
                List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();

                foreach (var uniqueID in distinctUniqIds)
                {
                    List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(x => x.UniqId == uniqueID).ToList();
                    QarsafariGrouped qarsafariGrouped = new QarsafariGrouped();



                    Qarsafari qarsafariExcel = qarsafaris.FirstOrDefault(m => m.IsUniqLiterNull == "true");


                    #region ახალი გადინამიურებული დაგრუპვა (დარჩენილია ფორმულების გაკეთება)
                    //if (qarsafariExcel != null)
                    //{
                    //    foreach (var columnName in geographicDynamicDbContext.ColumnNames)
                    //    {
                    //        //Get cell type
                    //        if (columnName.GroupMethod == "MAX")
                    //        {
                    //            if (qarsafariExcel != null)
                    //            {
                    //                PropertyInfo propertyInfoQarsafariGrouped = typeof(QarsafariGrouped).GetProperty(columnName.Sqlname);
                    //                PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(columnName.Sqlname);
                    //                object propertyValueQarsafari = propertyInfoQarsafari.GetValue(qarsafariExcel);

                    //                // თუ ცარიელია ვამოწმებთ და ვწერთ null -ს 
                    //                if (propertyValueQarsafari != null)
                    //                {

                    //                    //აქ სადაც double ია მოაქვს System.Single ამიტომ ვამოწმებთ

                    //                    if (propertyValueQarsafari.GetType() == typeof(System.Single) || propertyValueQarsafari.GetType() == typeof(System.Double))
                    //                    {
                    //                        Double? doubleValue = Convert.ToDouble(propertyValueQarsafari);

                    //                        propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, doubleValue, null);
                    //                    }
                    //                    else // სხვა შემთხვევაში არის string ან date ან bit
                    //                    {

                    //                        propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, propertyInfoQarsafari.GetValue(qarsafariExcel, null), null);
                    //                    }
                    //                }
                    //                else
                    //                {
                    //                    propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, null, null);
                    //                }

                    //                //if (propertyInfo != null)
                    //                //{

                    //                //    propertyInfo.SetValue(qarsafariGrouped, qarsafaris);
                    //                //}
                    //            }
                    //        }
                    //        else if (columnName.GroupMethod == "SUBSTRING")
                    //        {
                    //            if (qarsafariExcel != null)
                    //            {

                    //                PropertyInfo propertyInfoQarsafariGrouped = typeof(QarsafariGrouped).GetProperty(columnName.Sqlname);
                    //                PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(columnName.Sqlname);

                    //                var storeStringValue = "";
                    //                foreach (var item in qarsafaris)
                    //                {
                    //                    object propertyValueQarsfari = propertyInfoQarsafari.GetValue(item);

                    //                    if (propertyValueQarsfari != null)
                    //                    {

                    //                        if (propertyValueQarsfari.GetType() == typeof(System.String))
                    //                        {
                    //                            storeStringValue += string.Concat(propertyValueQarsfari, "/"); ;
                    //                            propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, storeStringValue.TrimEnd('/'), null);
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }


                    //        else if (columnName.GroupMethod == "SUM")
                    //        {
                    //            if (qarsafariExcel != null)
                    //            {

                    //                PropertyInfo propertyInfoQarsafariGrouped = typeof(QarsafariGrouped).GetProperty(columnName.Sqlname);
                    //                PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(columnName.Sqlname);

                    //                Double? storedValue = 0;
                    //                foreach (var item in qarsafaris)
                    //                {
                    //                    object propertyValueQarsafari = propertyInfoQarsafari.GetValue(item);
                    //                    if (propertyValueQarsafari != null)
                    //                    {
                    //                        if ((propertyValueQarsafari.GetType() == typeof(System.Single)) || (propertyValueQarsafari.GetType() == typeof(System.Double)))
                    //                        {
                    //                            storedValue += (double)propertyValueQarsafari;

                    //                            propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, storedValue, null);
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }
                    //    }
                    //    geographicDynamicDbContext.QarsafariGroupeds.Add(qarsafariGrouped);
                    //    geographicDynamicDbContext.SaveChanges();
                    //}

                    #endregion



                    #region ჩვეულებრივი ხელით
                    if (qarsafariExcel != null)
                    {
                        //Double? woodyplantqunatity = 0;
                        Double? merqmcenarisPr = 0;
                        Double? romgitxariGadanomrili = 0;
                        string mcenarisSaxeobebi = "";
                        string mcenarisSaxeobebiCorrected = "";
                        //double? InGoodCondition = 100;


                        //კარგ მდგომარეობაში ჩასაწერი ველისთვის 
                        Double? sumWoodyPlantQuantitymultiplyChoppedDown = 0;
                        Double? sumWoodyPlantQuantity = 0;
                        Double? sumWoodyPlantQuantitymultiplyRampike = 0;

                        //საშუალო ხმოვანებაში ჩასაწერი ველისთვის 
                        double? speciesMidAge = 0;
                        foreach (var qarsafari in qarsafaris)
                        {

                            //woodyplantqunatity += qarsafari.WoodyPlantQuantity;
                            // ამაში ამოვაგდეთ გაჩეხილი (chopped_down)
                            merqmcenarisPr += qarsafari.ChoppedDownQuantity != null ? (((qarsafari.WoodyPlantQuantity - qarsafari.ChoppedDownQuantity) * qarsafari.VarjisFarti) / qarsafariExcel.LandAreaSqM) * 100 : ((qarsafari.WoodyPlantQuantity * qarsafari.VarjisFarti) / qarsafariExcel.LandAreaSqM) * 100;
                            //merqmcenarisPr += ((qarsafari.WoodyPlantQuantity * qarsafari.VarjisFarti) / qarsafariExcel.LandAreaSqM) * 100;
                            romgitxariGadanomrili += qarsafari.WoodyPlantQuantity * qarsafari.VarjisFarti;
                            //mcenarisSaxeobebi += string.Concat(qarsafari.WoodyPlantSpecies, "/ "); // აქ იყო  "/ "
                            if (!string.IsNullOrEmpty(qarsafari.WoodyPlantSpecies))// კეთდება შემოწმება იმისთვის რომ გავიგოთ სლექში ჩაიწეროს თუ არა ველში 
                            {
                                mcenarisSaxeobebi += qarsafari.WoodyPlantSpecies + "/";
                            }

                            //InGoodCondition -= ((qarsafari.WoodyPlantQuantity * qarsafari.ChoppedDown) / 100) / qarsafari.WoodyPlantQuantity;

                            // კარგ მდომარეობაში ჩასაწერი 
                            // ვითვლით SUM([Woody_plant_quantity] * [chopped_down] / 100)
                            sumWoodyPlantQuantitymultiplyChoppedDown += ((qarsafari.WoodyPlantQuantity == null ? 0 : qarsafari.WoodyPlantQuantity) * (qarsafari.ChoppedDown == null ? 0 : qarsafari.ChoppedDown)) / 100;
                            // ვითვლით მცენარეების რაოდენობის მთლიან ჯამს
                            sumWoodyPlantQuantity += qarsafari.WoodyPlantQuantity == null ? 0 : qarsafari.WoodyPlantQuantity;
                            // ვითვლით SUM([Woody_plant_quantity] * [rampike] / 100)
                            sumWoodyPlantQuantitymultiplyRampike += ((qarsafari.WoodyPlantQuantity == null ? 0 : qarsafari.WoodyPlantQuantity) * (qarsafari.Rampike == null ? 0 : qarsafari.Rampike)) / 100;

                            //საშუალო ხმოვანეის ჩასაწერი 
                            speciesMidAge += (qarsafari.SpeciesMediumAge * qarsafari.WoodyPlantQuantity);

                        }


                        qarsafariGrouped.UniqId = qarsafariExcel.UniqId;
                        qarsafariGrouped.LiterId = qarsafariExcel.LiterId;
                        qarsafariGrouped.PhotoN = qarsafariExcel.PhotoN;
                        qarsafariGrouped.Region = qarsafariExcel.Region;
                        qarsafariGrouped.Municipality = qarsafariExcel.Municipality;
                        qarsafariGrouped.AdmMun = qarsafariExcel.AdmMun;
                        qarsafariGrouped.CityTownVillage = qarsafariExcel.CityTownVillage;
                        qarsafariGrouped.LandAreaSqM = qarsafariExcel.LandAreaSqM;
                        qarsafariGrouped.LandAreaHa = Math.Round(Convert.ToDouble(qarsafariExcel.LandAreaHa), 1);
                        qarsafariGrouped.Shrubbery = qarsafariExcel.Shrubbery;
                        qarsafariGrouped.WoodyPlantPercent = Math.Round(Convert.ToDouble(merqmcenarisPr), 1);
                        qarsafariGrouped.WoodyPlantQuantity = sumWoodyPlantQuantity;
                        qarsafariGrouped.VarjisFarti = romgitxariGadanomrili;
                        // ვაშორებთ ბოლო "/" -ს
                        mcenarisSaxeobebiCorrected = mcenarisSaxeobebi.TrimEnd('/');
                        qarsafariGrouped.WoodyPlantSpecies = mcenarisSaxeobebiCorrected;
                        // კარგ მდგომარეობაში არის 100 - გამხმარი - გაჩეხილი
                        //თუ ხეები არ გვაქვს ჩაწეროს 0

                        qarsafariGrouped.InGoodCondition = sumWoodyPlantQuantity == 0 ? 0 : 100 - (sumWoodyPlantQuantity == 0 ? 0 : (Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyChoppedDown / sumWoodyPlantQuantity) * 100), 0) + Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyRampike / sumWoodyPlantQuantity) * 100), 0)));
                        qarsafariGrouped.ChoppedDown = sumWoodyPlantQuantity == 0 ? 0 : (Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyChoppedDown / sumWoodyPlantQuantity) * 100), 0));
                        qarsafariGrouped.Rampike = sumWoodyPlantQuantity == 0 ? 0 : Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyRampike / sumWoodyPlantQuantity) * 100), 0);
                        qarsafariGrouped.SpeciesMediumAge = RoundToNearest(Convert.ToDouble(speciesMidAge / sumWoodyPlantQuantity)); // Math.Round(Convert.ToDouble((((speciesMidAge / sumWoodyPlantQuantity / 5) * 0.5) * 10)), 0);
                        qarsafariGrouped.Note = qarsafariExcel.Note;
                        qarsafariGrouped.Company = qarsafariExcel.Company;
                        qarsafariGrouped.LandGisOperator = qarsafariExcel.LandGisOperator;
                        qarsafariGrouped.Date = qarsafariExcel.Date;
                        qarsafariGrouped.GisOperator = qarsafariExcel.GisOperator;
                        qarsafariGrouped.FieldOperator = qarsafariExcel.FieldOperator;
                        qarsafariGrouped.DaTe1 = qarsafariExcel.DaTe1;
                        qarsafariGrouped.OverlapCadCode = qarsafariExcel.OverlapCadCode;
                        qarsafariGrouped.Owner = qarsafariExcel.Owner;
                        qarsafariGrouped.Sakutreba = qarsafariExcel.Sakutreba;
                        qarsafariGrouped.LegalPerson = qarsafariExcel.LegalPerson;
                        qarsafariGrouped.UniqIdOld = qarsafariExcel.UniqIdOld;

                        //რაუნდები დაუწერე აქ, ზემოთ არ დაუწერო
                        //qarsafariGrouped.WoodyPlantPercent = merqmcenarisPr;


                        geographicDynamicDbContext.QarsafariGroupeds.Add(qarsafariGrouped);
                        geographicDynamicDbContext.SaveChanges();
                    }
                    #endregion
                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა დაგრუპვა ქარსაფარის ცხრილის"
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ქარსაფარის ცხრილის დაგრუპვის დროს" + ex.Message
                };
            }
        }
        // qarsafariGrouped ცხრილის UID ველის შევსება ლიტერით და უნიკაიდით 
        public Result<bool> UIDReplaceQarsafariGrouped()
        {

            try
            {

                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds.ToList();
                foreach (var item in qarsafariGroupeds)
                {
                    item.Uid = String.Concat(item.LiterId, item.UniqId);
                    GeographicDynamicDbContext.SaveChanges();

                }
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }

            catch (Exception ex)
            {

                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა UID replace-ის დროს" + ex.Message
                };
            }
        }
        #region WindbreakMDB SQL-ში ჩაწერა QarsafariGroupded-ან
        //ფუნქცია იმისთვის რომ Access ფაილში ჩაიწეროს QarsafariGroupded-ან 
        /// ////// ეს ფუქნცია უბრალოდ SQL tablshi ყრის ქარსაფარიდან და მერე ვეღარ ვიყენებთ ჯერჯერობით ვაკომენტარებ სამომავლოდ შეიძლება რამეში გამოვიყენოთ 
        //ეს არ წერს Mdb ში არაფერს !!!!!!!!!!!!!!!!
        //public Result<bool> UPDTFromExcelToAccess(string AccessShitName)
        //{
        //    try
        //    {

        //        var GeographicDynamicDbContext = new GeographicDynamicDbContext();

        //        List<WindbreakMdb> windbreakMdbs = GeographicDynamicDbContext.WindbreakMdbs.ToList();
        //        List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds.ToList();

        //        if (!string.IsNullOrEmpty(AccessShitName))
        //        {
        //            foreach (var item in windbreakMdbs)
        //            {
        //                //WindbreakMdb access = GeographicDynamicDbContext.WindbreakMdbs.FirstOrDefault(x => x.LiterId == excel.LiterId && x.UniqId == excel.UniqId);
        //                QarsafariGrouped ExcelGrouped = GeographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.Uid == item.Uid);
        //                if (ExcelGrouped != null)
        //                {
        //                    item.PhotoN = ExcelGrouped.PhotoN;
        //                    item.Shrubbery = Convert.ToDouble(ExcelGrouped.Shrubbery);
        //                    item.WoodyPlantPercent = Convert.ToString(ExcelGrouped.WoodyPlantPercent);
        //                    item.WoodyPlantQuantity = Convert.ToDouble(ExcelGrouped.WoodyPlantQuantity);
        //                    item.WoodyPlantSpecies = ExcelGrouped.WoodyPlantSpecies;
        //                    item.InGoodCondition = Convert.ToDouble(ExcelGrouped.InGoodCondition);
        //                    item.ChoppedDown = Convert.ToDouble(ExcelGrouped.ChoppedDown);
        //                    item.Rampike = Convert.ToDouble(ExcelGrouped.Rampike);
        //                    item.SpeciesMediumAge = Convert.ToDouble(ExcelGrouped.SpeciesMediumAge);
        //                    item.Company = ExcelGrouped.Company;
        //                    item.FieldOperator = ExcelGrouped.FieldOperator;
        //                    item.UniqId = (float?)ExcelGrouped.UniqId;

        //                    GeographicDynamicDbContext.SaveChanges();

        //                }
        //            }
        //        }


        //        return new Result<bool>
        //        {
        //            Success = true,
        //            StatusCode = System.Net.HttpStatusCode.OK
        //        };

        //    }

        //    catch (Exception ex)
        //    {

        //        return new Result<bool>
        //        {
        //            Success = false,
        //            StatusCode = System.Net.HttpStatusCode.BadGateway,
        //            Message = "მოხდა შეცდომა Excel-ცხრილიდან Access-ცხრილში გადაწერის დროს" + ex.Message
        //        };
        //    }
        //}
        #endregion



        //////////// ფუნქცია კითხულობს SQL-ბაზას კონკრეტულად qarsafariGroupeds-ს და წერს დათვლილ საჭირო მონაცემებს თვითონ access ფაილში 
        #region ChatGPT + gios-ს ნახლაფორთალი რომელიც იღებს qarsafariGroupds და საჭირო ველები რომლებიც გვჭირდება access ფაილში იწერება იქ ჯერ სორტირდება და შემდეგ იყრება 
        #region მონახაზი მარა მაინც იყოს რა იცი რაში დაგჭირდეს კაცს Access ფაილში მონაცემების ჩაწერის
        //public Result<bool> UpdateFromQarsafariGroupedToAccessFile(string AccessShitName, string AccessFilePath)
        //{
        //    try
        //    {

        //        var GeographicDynamicDbContext = new GeographicDynamicDbContext();
        //        var uniqid = GeographicDynamicDbContext.ColumnNames.FirstOrDefault(m => m.Sqlname == "UniqId").AccessName;
        //        var literid = GeographicDynamicDbContext.ColumnNames.FirstOrDefault(m => m.Sqlname == "LiterId").AccessName;
        //        //var filterAccess = columnNameDTO.AccessName;
        //        var AccessFileAddress = AccessFilePath;

        //        // Sort qarsafariGroupeds by UniqId
        //        List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds
        //            .OrderBy(x => x.UniqId)
        //            .ToList();

        //        #region  OleDbConnection for Access

        //        //OleDbConnection

        //        //string connectionString = @"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\Users\gioch\OneDrive\Desktop\GEOGraphics\test.accdb";
        //        //string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\gioch\\OneDrive\\Desktop\\GEOGraphics\Dedoplistskaro.mdb";
        //        string connectionString = "";
        //        if (Path.GetExtension(AccessFileAddress).ToLower().Trim() == ".mdb" && Environment.Is64BitOperatingSystem == false)
        //        {
        //            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AccessFileAddress;
        //            connectionString = "Provider=Microsoft.Jet.OLEDBMicrosoft.Jet.OLEDB.4.0;Data Source=" + AccessFileAddress + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
        //        }
        //        else
        //        {
        //            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AccessFileAddress;
        //        }

        //        #endregion

        //        using (OleDbConnection connection = new OleDbConnection(connectionString))
        //        {
        //            connection.Open();

        //            // Construct the update command
        //            string updateCommand = $"UPDATE [{AccessShitName}] SET PhotoN = @Photo_N WHERE {uniqid} = @UNIQ_ID AND {literid} = @Liter_ID";
        //            OleDbCommand command = new OleDbCommand(updateCommand, connection);

        //            // Iterate through qarsafariGroupeds and update Access file
        //            foreach (var item in qarsafariGroupeds)
        //            {
        //                command.Parameters.Clear();
        //                command.Parameters.AddWithValue("@Photo_N", item.PhotoN);
        //                command.Parameters.AddWithValue("@UNIQ_ID", item.UniqId);
        //                command.Parameters.AddWithValue("@Liter_ID", item.LiterId);

        //                command.ExecuteNonQuery();
        //            }

        //            connection.Close();
        //        }


        //        return new Result<bool>
        //        {
        //            Success = true,
        //            StatusCode = System.Net.HttpStatusCode.OK
        //        };
        //    }
        //    catch (Exception ex)
        //    {
        //        return new Result<bool>
        //        {
        //            Success = false,
        //            StatusCode = System.Net.HttpStatusCode.BadGateway,
        //            Message = "An error occurred while updating the Access table from the Excel data: " + ex.Message
        //        };
        //    }
        //}
        #endregion


        ///////// ეს ფუნქცია ამატებს access ფალში გადანომრილ ინფორმაციას 
        public Result<bool> UpdateFromQarsafariGroupedToAccessFile(string AccessSheetName, string AccessFilePath)
        {
            try
            {
                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                var uniqid = GeographicDynamicDbContext.ColumnNames.FirstOrDefault(m => m.Sqlname == "UniqId").AccessName;
                var literid = GeographicDynamicDbContext.ColumnNames.FirstOrDefault(m => m.Sqlname == "LiterId").AccessName;

                //// Sort qarsafariGroupeds by UniqId and LiterId
                //List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds
                //    .OrderBy(x => x.UniqId)
                //    //.ThenBy(x => x.LiterId)
                //    .ToList();

                // Connection string for Access
                string connectionString = "";
                if (Path.GetExtension(AccessFilePath).ToLower().Trim() == ".mdb" && !Environment.Is64BitOperatingSystem)
                {
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AccessFilePath;
                }
                else
                {
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AccessFilePath;
                }

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();



                    // Add new column if it does not exist
                    string newColumnName = "Uniq_ID_NEW_Gadanomrili";
                    try
                    {
                        string alterTableQuery = $"ALTER TABLE [{AccessSheetName}] ADD COLUMN {newColumnName} DOUBLE";
                        OleDbCommand alterCmd = new OleDbCommand(alterTableQuery, connection);
                        alterCmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        if (!ex.Message.Contains("duplicate") && !ex.Message.Contains("already exists"))
                        {
                            throw;
                        }
                    }


                    // Sort Access table by UniqId and LiterId
                    //string sortCommand = $"SELECT * FROM [{AccessSheetName}] ORDER BY {uniqid}, {literid}";
                    string query = $"SELECT * FROM [{AccessSheetName}]";
                    OleDbCommand sortCmd = new OleDbCommand(query, connection);
                    System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(sortCmd);
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);

                    // Update PhotoN values in Access table
                    foreach (var item in GeographicDynamicDbContext.QarsafariGroupeds)
                    {
                        // Find corresponding row in Access table
                        System.Data.DataRow[] rows = dataTable.Select($"{uniqid} = '{item.UniqIdOld}' AND {literid} = {item.LiterId}");

                        if (rows.Length > 0)
                        {

                            foreach (var row in rows) // update all matches (in case there are more than one)
                            {
                                row["Photo_N"] = item.PhotoN;
                                row["shrubbery"] = item.Shrubbery;
                                row["Woody_plant_percent"] = item.WoodyPlantPercent;
                                row["Woody_plant_quantity"] = item.WoodyPlantQuantity;
                                row["woody_plant_species"] = item.WoodyPlantSpecies;
                                row["In_good_condition"] = item.InGoodCondition;
                                row["chopped_down"] = item.ChoppedDown;
                                row["rampike"] = item.Rampike;
                                row["species_medium_age"] = item.SpeciesMediumAge;
                                row["Company"] = item.Company;
                                row["Field_Operator"] = item.FieldOperator;
                                row["Date_"] = item.Date;
                                row[newColumnName] = item.UniqId;
                                row["UNIQ_ID"] = item.UniqId;
                                row["UNIQ_ID_OLD"] = item.UniqIdOld;
                            }
                            //rows[0]["Photo_N"] = item.PhotoN;
                            //rows[0]["shrubbery"] = item.Shrubbery;
                            //rows[0]["Woody_plant_percent"] = item.WoodyPlantPercent;
                            //rows[0]["Woody_plant_quantity"] = item.WoodyPlantQuantity;
                            //rows[0]["woody_plant_species"] = item.WoodyPlantSpecies;
                            //rows[0]["In_good_condition"] = item.InGoodCondition;
                            //rows[0]["chopped_down"] = item.ChoppedDown;
                            //rows[0]["rampike"] = item.Rampike;
                            //rows[0]["species_medium_age"] = item.SpeciesMediumAge;
                            //rows[0]["Company"] = item.Company;
                            //rows[0]["Field_Operator"] = item.FieldOperator;
                            //rows[0]["Date_"] = item.Date;
                            //rows[0][newColumnName] = item.UniqId;
                            //rows[0]["UNIQ_ID"] = item.UniqId;
                            //rows[0]["UNIQ_ID_OLD"] = item.UniqIdOld;
                        }
                    }

                    // Update Access table with modified DataTable
                    System.Data.OleDb.OleDbCommandBuilder builder = new System.Data.OleDb.OleDbCommandBuilder(adapter);
                    adapter.UpdateCommand = builder.GetUpdateCommand();
                    adapter.Update(dataTable);

                    connection.Close();
                }


                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {

                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.InternalServerError,
                    Message = "An error occurred while updating the Access table: " + ex.Message
                };
            }
        }



        #endregion





        // ფუნქცია ყრის მონაცემებს gadanomriliFotoebi-დან qarsafariGroupded-ში 
        public Result<bool> GadanomriliFotoebiToQarsafariGrouped()
        {
            try
            {

                var GeographicDynamicDbContext = new GeographicDynamicDbContext();

                List<GadanomriliFotoebi> FotoList = GeographicDynamicDbContext.GadanomriliFotoebis.ToList();
                List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds.ToList();

                {
                    foreach (var item in FotoList)
                    {
                        QarsafariGrouped ExcelGrouped = GeographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.LiterId == item.LiterId && x.UniqId == Convert.ToDouble(item.UniqId));
                        if (ExcelGrouped != null)
                        {
                            ExcelGrouped.PhotoN = item.PhotoN;
                            ExcelGrouped.Date = item.PhotoDate;

                            GeographicDynamicDbContext.SaveChanges();

                        }
                    }


                    return new Result<bool>
                    {
                        Success = true,
                        StatusCode = System.Net.HttpStatusCode.OK
                    };

                }
            }

            catch (Exception ex)
            {

                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა GadanomriliPhotoebi-ცხრილიდან QarsafariGrouped-ცხრილში გადაწერის დროს" + ex.Message
                };
            }
        }
        // ექსელში ჩაწერა
        public Result<bool> WriteToExcel(List<Qarsafari> qarsafaris, string ExcelDestinationPath, string ExcelName)
        {
            var GeographicDynamicDbContext = new GeographicDynamicDbContext();

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("დიდი-ექსელი");

                    // ამით ივსება სათაურების ველები 
                    worksheet.Cell(1, "A").Value = "UNIQ_ID";
                    worksheet.Cell(1, "B").Value = "Liter_ID";
                    worksheet.Cell(1, "C").Value = "Photo_N";
                    worksheet.Column("C").Style.NumberFormat.Format = "@";
                    worksheet.Cell(1, "D").Value = "Region";
                    worksheet.Cell(1, "E").Value = "Municipality";
                    worksheet.Cell(1, "F").Value = "Adm_Mun";
                    worksheet.Cell(1, "G").Value = "City_Town_Village";
                    worksheet.Cell(1, "H").Value = "Land_Area_Sq_M";
                    worksheet.Cell(1, "I").Value = "Land_Area_Ha";
                    worksheet.Cell(1, "J").Value = "Shrubbery";
                    worksheet.Cell(1, "K").Value = "Woody_Plant_Percent";
                    worksheet.Cell(1, "L").Value = "Woody_Plant_Quantity";
                    worksheet.Cell(1, "M").Value = "Woody_Plant_Spices";
                    worksheet.Cell(1, "N").Value = "VarjisFarti";
                    worksheet.Cell(1, "O").Value = "In_Good_Condition";
                    worksheet.Cell(1, "P").Value = "Chopped_down";
                    worksheet.Cell(1, "Q").Value = "Rampike";
                    worksheet.Cell(1, "R").Value = "Spices_Medium_Age";
                    worksheet.Cell(1, "S").Value = "Note_";
                    worksheet.Cell(1, "T").Value = "Company";
                    worksheet.Cell(1, "U").Value = "Field_Operator";
                    worksheet.Cell(1, "V").Value = "Date_";
                    worksheet.Column("V").Style.NumberFormat.Format = "@"; //ფორმატტდება დეითის ველის ტიპი ტექსტად
                    worksheet.Cell(1, "W").Value = "Gis_Operator";
                    worksheet.Cell(1, "X").Value = "DaTe_1";
                    worksheet.Column("X").Style.NumberFormat.Format = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                    worksheet.Cell(1, "Y").Value = "Overlap_CAD_CODE";
                    worksheet.Cell(1, "Z").Value = "Owner";
                    worksheet.Cell(1, "AA").Value = "Legal_person";
                    worksheet.Cell(1, "AB").Value = "Owners";
                    worksheet.Cell(1, "AC").Value = "Land_Field_Operator";
                    worksheet.Cell(1, "AD").Value = "Note1";
                    worksheet.Cell(1, "AE").Value = "Date_2";
                    worksheet.Cell(1, "AF").Value = "Land_Gis_Operator";
                    worksheet.Cell(1, "AG").Value = "Note1_1";
                    worksheet.Cell(1, "AH").Value = "Date_3";
                    worksheet.Cell(1, "AI").Value = "CAD_COD";
                    worksheet.Cell(1, "AJ").Value = "UNIQ_ID_OLD";
                    worksheet.Cell(1, "AK").Value = "UNIQ_ID_NEW";
                    worksheet.Cell(1, "AL").Value = "UID";
                    worksheet.Cell(1, "AM").Value = "ID";

                    // ამ ციკლით ივსება Rows სათაურების ქვეშ 
                    for (int r = 0; r < qarsafaris.Count(); r++) //r stands for ExcelRow and c for ExcelColumn
                    {
                        QarsafariGrouped qarsafariGrouped = GeographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.UniqId == qarsafaris[r].UniqId);
                        int rowNum = r + 2;

                        if (qarsafaris[r].IsUniqLiterNull == "true")
                        {
                            worksheet.Cell(rowNum, "A").Value = qarsafaris[r].UniqId;
                            worksheet.Cell(rowNum, "B").Value = qarsafaris[r].LiterId;
                            worksheet.Cell(rowNum, "AJ").Value = qarsafaris[r].UniqIdOld;
                            worksheet.Cell(rowNum, "K").Value = Math.Round(Convert.ToDouble(qarsafariGrouped?.WoodyPlantPercent ?? 0), 1);
                            worksheet.Cell(rowNum, "C").Value = qarsafariGrouped?.PhotoN; // ფოტოები მოდის დაგრუპულიდან 
                            worksheet.Cell(rowNum, "D").Value = qarsafaris[r].Region;
                            worksheet.Cell(rowNum, "E").Value = qarsafaris[r].Municipality;
                            worksheet.Cell(rowNum, "F").Value = qarsafaris[r].AdmMun;
                            worksheet.Cell(rowNum, "G").Value = qarsafaris[r].CityTownVillage;
                            worksheet.Cell(rowNum, "H").Value = qarsafaris[r].LandAreaSqM;
                            worksheet.Cell(rowNum, "I").Value = qarsafaris[r].LandAreaHa;
                            worksheet.Cell(rowNum, "V").Value = qarsafariGrouped?.Date;
                            worksheet.Cell(rowNum, "Z").Value = qarsafaris[r].Owner;
                        }

                        worksheet.Cell(rowNum, "J").Value = qarsafaris[r].Shrubbery;
                        worksheet.Cell(rowNum, "L").Value = qarsafaris[r].WoodyPlantQuantity;
                        worksheet.Cell(rowNum, "M").Value = qarsafaris[r].WoodyPlantSpecies;
                        worksheet.Cell(rowNum, "N").Value = qarsafaris[r].VarjisFarti;

                        // მოწმდება თუ სადმე sumofGoodChoppedRampike განსხვავდება 0-ს ან 100-ს იდეაში 0.1 ან მეტია ან ნაკლები და მაგის მიხედვით 
                        // ხორციელდება გამოკლება ან მიმატება 0.1-ის 
                        var sumofGoodChoppedRampike = Math.Round(Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition ?? 0), 1)
                            + Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown ?? 0), 1)
                            + Math.Round(Convert.ToDouble(qarsafaris[r].Rampike ?? 0), 1), 1);

                        switch (sumofGoodChoppedRampike)
                        {
                            case 99.9:
                                worksheet.Cell(rowNum, "O").Value = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition ?? 0), 1) + 0.1;
                                worksheet.Cell(rowNum, "P").Value = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown ?? 0), 1);
                                worksheet.Cell(rowNum, "Q").Value = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike ?? 0), 1);
                                break;
                            case 100.1:
                                if (qarsafaris[r].InGoodCondition > 0)
                                {
                                    worksheet.Cell(rowNum, "O").Value = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition ?? 0), 1) - 0.1;
                                    worksheet.Cell(rowNum, "P").Value = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown ?? 0), 1);
                                    worksheet.Cell(rowNum, "Q").Value = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike ?? 0), 1);
                                }
                                else if (qarsafaris[r].ChoppedDown > 0)
                                {
                                    worksheet.Cell(rowNum, "O").Value = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition ?? 0), 1);
                                    worksheet.Cell(rowNum, "P").Value = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown ?? 0), 1) - 0.1;
                                    worksheet.Cell(rowNum, "Q").Value = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike ?? 0), 1);
                                }
                                else
                                {
                                    worksheet.Cell(rowNum, "O").Value = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition ?? 0), 1);
                                    worksheet.Cell(rowNum, "P").Value = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown ?? 0), 1);
                                    worksheet.Cell(rowNum, "Q").Value = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike ?? 0), 1) - 0.1;
                                }
                                break;
                            default:
                                worksheet.Cell(rowNum, "O").Value = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition ?? 0), 1);
                                worksheet.Cell(rowNum, "P").Value = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown ?? 0), 1);
                                worksheet.Cell(rowNum, "Q").Value = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike ?? 0), 1);
                                break;
                        }

                        worksheet.Cell(rowNum, "R").Value = qarsafaris[r].SpeciesMediumAge;
                        worksheet.Cell(rowNum, "S").Value = qarsafaris[r].Note;
                        worksheet.Cell(rowNum, "T").Value = qarsafaris[r].Company;
                        worksheet.Cell(rowNum, "U").Value = qarsafaris[r].FieldOperator;
                        worksheet.Cell(rowNum, "W").Value = qarsafaris[r].GisOperator;
                        worksheet.Cell(rowNum, "X").Value = qarsafaris[r].DaTe1;
                        worksheet.Cell(rowNum, "Y").Value = qarsafaris[r].OverlapCadCode;
                        worksheet.Cell(rowNum, "AA").Value = qarsafaris[r].LegalPerson;
                        worksheet.Cell(rowNum, "AB").Value = qarsafaris[r].Owners;
                        worksheet.Cell(rowNum, "AC").Value = qarsafaris[r].LandFieldOperator;
                        worksheet.Cell(rowNum, "AD").Value = qarsafaris[r].Note1;
                        worksheet.Cell(rowNum, "AE").Value = qarsafaris[r].Date2;
                        worksheet.Cell(rowNum, "AF").Value = qarsafaris[r].LandGisOperator;
                        worksheet.Cell(rowNum, "AG").Value = qarsafaris[r].Note11;
                        worksheet.Cell(rowNum, "AH").Value = qarsafaris[r].Date3;
                        worksheet.Cell(rowNum, "AI").Value = qarsafaris[r].CadCod;
                        worksheet.Cell(rowNum, "AK").Value = qarsafaris[r].UniqIdNew;
                        worksheet.Cell(rowNum, "AL").Value = qarsafaris[r].Uid;
                        worksheet.Cell(rowNum, "AM").Value = qarsafaris[r].Id;
                    }

                    // Create result folder if it doesn't exist
                    string resultFolderPath = Path.Combine(ExcelDestinationPath, "result");
                    if (!Directory.Exists(resultFolderPath))
                    {
                        Directory.CreateDirectory(resultFolderPath);
                    }

                    workbook.SaveAs(Path.Combine(resultFolderPath, $"{ExcelName}.xlsx"));
                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა SQL-დან ახალ Execl-ში გადაწერის დროს" + ex.Message
                };
            }
        }



        public class storedMDB
        {
            public string Uniq_Id_MDB { get; set; }
            public string Uniq_ID_gadanomrili { get; set; }
            public string Litter_Id { get; set; }
        }
        List<storedMDB> excelDataList = new List<storedMDB>();

        //ფუნქცია კითხულობს ბაზას და ქმნის ახალ ექსელის ფაილს რომ ჩაიწეროს მონაცემები მხოლოდ დაგრუპულისთვის 
        public Result<bool> WriteToExcelGrouped(List<QarsafariGrouped> qarsafariGroupeds, string ExcelDestinationPath, string ExcelName, string AccessPath, string AccessSheetName)
        {

            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext(); //უკავშირდება კონტექსტს რომ გაიგოს ცხრილები SQL-დან 

            // Dictionary to store the mapping from the Access file
            Dictionary<string, string> accessData = new Dictionary<string, string>();

            try
            {
                // Open and read the Access file


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



                using (OleDbConnection accessConn = new OleDbConnection(connectionString))
                {
                    accessConn.Open();
                    string query = $"SELECT * FROM [{AccessSheetName}]";
                    using (OleDbCommand cmd = new OleDbCommand(query, accessConn))
                    {
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                int uniqIdNewIndex = reader.GetOrdinal("Uniq_ID_NEW_Gadanomrili");
                                int uniqIdIndex = reader.GetOrdinal("UNIQ_ID_OLD");
                                int litterId = reader.GetOrdinal("Liter_ID");

                                while (reader.Read())
                                {
                                    string uniqIdNewValue = reader.GetValue(uniqIdNewIndex).ToString();
                                    string uniqIdValue = reader.GetValue(uniqIdIndex).ToString();
                                    string LitterIdValue = reader.GetValue(litterId).ToString();


                                    if (!accessData.ContainsKey(uniqIdNewValue))
                                    {
                                        accessData[uniqIdNewValue] = uniqIdValue;
                                    }
                                    excelDataList.Add(new storedMDB
                                    {
                                        Uniq_Id_MDB = uniqIdValue,
                                        Uniq_ID_gadanomrili = uniqIdNewValue,
                                        Litter_Id = LitterIdValue
                                    });

                                }
                            }
                        }
                    }
                }






                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("პატარა-ექსელი");

                    // ამ კოდის ფრაგმენტებში ივსება სათაურის ველები სხვა სიტყვებით რომ ვთქვათ პირველ row-ში იწერება მნიშვნელობები რამდენი სვეტიც გვაქ (column) 
                    worksheet.Cell(1, "A").Value = "UNIQ_ID";
                    worksheet.Cell(1, "B").Value = "Liter_ID";
                    worksheet.Cell(1, "C").Value = "Photo_N";
                    worksheet.Column("C").Style.NumberFormat.Format = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                    worksheet.Cell(1, "D").Value = "REGION";
                    worksheet.Cell(1, "E").Value = "Municipality";
                    worksheet.Cell(1, "F").Value = "Adm_Mun";
                    worksheet.Cell(1, "G").Value = "City_Town_Village";
                    worksheet.Cell(1, "H").Value = "Land_Area_Sq_m";
                    worksheet.Cell(1, "I").Value = "Land_Area_Ha";
                    worksheet.Cell(1, "J").Value = "shrubbery";
                    worksheet.Cell(1, "K").Value = "Woody_plant_percent";
                    worksheet.Cell(1, "L").Value = "Woody_plant_quantity";
                    worksheet.Cell(1, "N").Value = "woody_plant_species";
                    worksheet.Cell(1, "O").Value = "In_good_condition";
                    worksheet.Cell(1, "P").Value = "chopped_down";
                    worksheet.Cell(1, "Q").Value = "rampike";
                    worksheet.Cell(1, "R").Value = "species_medium_age";
                    worksheet.Cell(1, "S").Value = "Note_";
                    worksheet.Cell(1, "T").Value = "Company";
                    worksheet.Cell(1, "U").Value = "Field_Operator";
                    worksheet.Cell(1, "V").Value = "Date_";
                    worksheet.Column("V").Style.NumberFormat.Format = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                    worksheet.Cell(1, "W").Value = "Gis_Operator";
                    worksheet.Cell(1, "X").Value = "DaTe_1";
                    worksheet.Column("X").Style.NumberFormat.Format = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                    worksheet.Cell(1, "Y").Value = "Overlap_CAD_CODE";
                    worksheet.Cell(1, "Z").Value = "Owner";
                    worksheet.Cell(1, "AA").Value = "Legal_person";
                    worksheet.Cell(1, "AB").Value = "Owners";
                    worksheet.Cell(1, "AC").Value = "Land_Field_Operator";
                    worksheet.Cell(1, "AD").Value = "Note1";
                    worksheet.Cell(1, "AE").Value = "Date_2";
                    worksheet.Cell(1, "AF").Value = "Land_Gis_Operator";
                    worksheet.Cell(1, "AG").Value = "Note1_1";
                    worksheet.Cell(1, "AH").Value = "Date_3";
                    worksheet.Cell(1, "AI").Value = "CAD_COD";
                    worksheet.Cell(1, "AJ").Value = "UNIQ_ID_OLD";
                    worksheet.Cell(1, "AK").Value = "UNIQ_ID_NEW";
                    worksheet.Cell(1, "AL").Value = "UID";
                    worksheet.Cell(1, "AM").Value = "ID";
                    worksheet.Cell(1, "AN").Value = "Uniq_ID_MDB";

                    for (var r = 0; r < qarsafariGroupeds.Count(); r++) // კეთდება ციკლი იმისთვის რო დაიაროს სათითაო ველი და ჩაიწეროს ექსელის შიტში 
                    {
                        int rowNum = r + 2;

                        worksheet.Cell(rowNum, "A").Value = qarsafariGroupeds[r].UniqId;
                        worksheet.Cell(rowNum, "B").Value = qarsafariGroupeds[r].LiterId;
                        worksheet.Cell(rowNum, "C").Value = qarsafariGroupeds[r].PhotoN;
                        worksheet.Cell(rowNum, "D").Value = qarsafariGroupeds[r].Region;
                        worksheet.Cell(rowNum, "E").Value = qarsafariGroupeds[r].Municipality;
                        worksheet.Cell(rowNum, "F").Value = qarsafariGroupeds[r].AdmMun;
                        worksheet.Cell(rowNum, "G").Value = qarsafariGroupeds[r].CityTownVillage;
                        worksheet.Cell(rowNum, "H").Value = qarsafariGroupeds[r].LandAreaSqM;
                        worksheet.Cell(rowNum, "I").Value = qarsafariGroupeds[r].LandAreaHa;
                        worksheet.Cell(rowNum, "J").Value = qarsafariGroupeds[r].Shrubbery;
                        worksheet.Cell(rowNum, "K").Value = qarsafariGroupeds[r].WoodyPlantPercent;
                        worksheet.Cell(rowNum, "L").Value = qarsafariGroupeds[r].WoodyPlantQuantity;
                        worksheet.Cell(rowNum, "N").Value = qarsafariGroupeds[r].WoodyPlantSpecies;

                        // სადაც სახეობა არ გვიწერია და ხეხილის რაოდენობა იქ იწერება კარგ მდომარეობაში 0 
                        if (qarsafariGroupeds[r].WoodyPlantQuantity == 0)
                        {
                            worksheet.Cell(rowNum, "O").Value = 0;
                            worksheet.Cell(rowNum, "P").Value = 0;
                            worksheet.Cell(rowNum, "Q").Value = 0;
                        }
                        else
                        {
                            worksheet.Cell(rowNum, "O").Value = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].InGoodCondition ?? 0), 1);
                            worksheet.Cell(rowNum, "P").Value = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].ChoppedDown ?? 0), 1);
                            worksheet.Cell(rowNum, "Q").Value = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].Rampike ?? 0), 1);
                        }
                        worksheet.Cell(rowNum, "R").Value = qarsafariGroupeds[r].SpeciesMediumAge;
                        worksheet.Cell(rowNum, "S").Value = qarsafariGroupeds[r].Note;
                        worksheet.Cell(rowNum, "T").Value = qarsafariGroupeds[r].Company;
                        worksheet.Cell(rowNum, "U").Value = qarsafariGroupeds[r].FieldOperator;
                        worksheet.Cell(rowNum, "V").Value = qarsafariGroupeds[r].Date;
                        worksheet.Cell(rowNum, "W").Value = qarsafariGroupeds[r].GisOperator;
                        worksheet.Cell(rowNum, "X").Value = qarsafariGroupeds[r].DaTe1;
                        worksheet.Cell(rowNum, "Y").Value = qarsafariGroupeds[r].OverlapCadCode;
                        worksheet.Cell(rowNum, "Z").Value = qarsafariGroupeds[r].Owner;
                        worksheet.Cell(rowNum, "AA").Value = qarsafariGroupeds[r].LegalPerson;
                        worksheet.Cell(rowNum, "AB").Value = qarsafariGroupeds[r].Owners;
                        worksheet.Cell(rowNum, "AC").Value = qarsafariGroupeds[r].LandFieldOperator;
                        worksheet.Cell(rowNum, "AD").Value = qarsafariGroupeds[r].Note1;
                        worksheet.Cell(rowNum, "AE").Value = qarsafariGroupeds[r].Date2;
                        worksheet.Cell(rowNum, "AF").Value = qarsafariGroupeds[r].LandGisOperator;
                        worksheet.Cell(rowNum, "AG").Value = qarsafariGroupeds[r].Note11;
                        worksheet.Cell(rowNum, "AH").Value = qarsafariGroupeds[r].Date3;
                        worksheet.Cell(rowNum, "AI").Value = qarsafariGroupeds[r].CadCod;
                        worksheet.Cell(rowNum, "AJ").Value = qarsafariGroupeds[r].UniqIdOld;
                        worksheet.Cell(rowNum, "AK").Value = qarsafariGroupeds[r].UniqIdNew;
                        worksheet.Cell(rowNum, "AL").Value = qarsafariGroupeds[r].Uid;
                        worksheet.Cell(rowNum, "AM").Value = qarsafariGroupeds[r].Id;

                        string uniqIdGadanomrili = qarsafariGroupeds[r].UniqId.ToString();
                        if (accessData.ContainsKey(uniqIdGadanomrili))
                        {
                            worksheet.Cell(rowNum, "AN").Value = accessData[uniqIdGadanomrili];
                        }
                    }

                    // Create result folder if it doesn't exist
                    string resultFolderPath = Path.Combine(ExcelDestinationPath, "result");
                    if (!Directory.Exists(resultFolderPath))
                    {
                        Directory.CreateDirectory(resultFolderPath);
                    }

                    workbook.SaveAs(Path.Combine(resultFolderPath, $"{ExcelName}.xlsx"));
                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ექსელში ჩაწერისას!" + ex.Message
                };
            }

        }



        //ფუნქცია კითხულობს ბაზას და ქმნის ახალ ექსელის ფაილს რომ ჩაიწეროს მონაცემები მხოლოდ დაგრუპულისთვის 
        public Result<bool> WriteToExcelRootOne(List<Qarsafari> newQarsafari, string ExcelDestinationPath)
        {

            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext(); //უკავშირდება კონტექსტს რომ გაიგოს ცხრილები SQL-დან 

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("პატარა-ექსელი");

                    // ამ კოდის ფრაგმენტებში ივსება სათაურის ველები სხვა სიტყვებით რომ ვთქვათ პირველ row-ში იწერება მნიშვნელობები რამდენი სვეტიც გვაქ (column) 
                    worksheet.Cell(1, "A").Value = "UNIQ_ID";
                    worksheet.Cell(1, "B").Value = "Liter_ID";
                    worksheet.Cell(1, "C").Value = "Photo_N";
                    worksheet.Cell(1, "D").Value = "REGION";
                    worksheet.Cell(1, "E").Value = "Municipality";
                    worksheet.Cell(1, "F").Value = "Adm_Mun";
                    worksheet.Cell(1, "G").Value = "City_Town_Village";
                    worksheet.Cell(1, "H").Value = "Land_Area_Sq_m";
                    worksheet.Cell(1, "I").Value = "Land_Area_Ha";
                    worksheet.Cell(1, "J").Value = "shrubbery";
                    worksheet.Cell(1, "K").Value = "Woody_plant_percent";
                    worksheet.Cell(1, "L").Value = "Woody_plant_quantity";
                    worksheet.Cell(1, "M").Value = "VarjisFarti";
                    worksheet.Cell(1, "N").Value = "woody_plant_species";
                    worksheet.Cell(1, "O").Value = "In_good_condition";
                    worksheet.Cell(1, "P").Value = "chopped_down";
                    worksheet.Cell(1, "Q").Value = "rampike";
                    worksheet.Cell(1, "R").Value = "species_medium_age";
                    worksheet.Cell(1, "S").Value = "Note_";
                    worksheet.Cell(1, "T").Value = "Company";
                    worksheet.Cell(1, "U").Value = "Field_Operator";
                    worksheet.Cell(1, "V").Value = "Date_";
                    worksheet.Column("V").Style.NumberFormat.Format = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                    worksheet.Cell(1, "W").Value = "Gis_Operator";
                    worksheet.Cell(1, "X").Value = "DaTe_1";
                    worksheet.Column("X").Style.NumberFormat.Format = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                    worksheet.Cell(1, "Y").Value = "Overlap_CAD_CODE";
                    worksheet.Cell(1, "Z").Value = "Owner";
                    worksheet.Cell(1, "AA").Value = "Legal_person";
                    worksheet.Cell(1, "AB").Value = "Owners";
                    worksheet.Cell(1, "AC").Value = "Land_Field_Operator";
                    worksheet.Cell(1, "AD").Value = "Note1";
                    worksheet.Cell(1, "AE").Value = "Date_2";
                    worksheet.Cell(1, "AF").Value = "Land_Gis_Operator";
                    worksheet.Cell(1, "AG").Value = "Note1_1";
                    worksheet.Cell(1, "AH").Value = "Date_3";
                    worksheet.Cell(1, "AI").Value = "CAD_COD";
                    worksheet.Cell(1, "AJ").Value = "UNIQ_ID_OLD";
                    worksheet.Cell(1, "AK").Value = "UNIQ_ID_NEW";
                    worksheet.Cell(1, "AL").Value = "UID";
                    worksheet.Cell(1, "AM").Value = "ID";
                    worksheet.Cell(1, "AN").Value = "Uniq_Id_NEW";

                    for (var r = 0; r < newQarsafari.Count(); r++)
                    {
                        QarsafariGrouped qarsafariGrouped = geographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.UniqIdOld == newQarsafari[r].UniqId);
                        int rowNum = r + 2;

                        worksheet.Cell(rowNum, "A").Value = newQarsafari[r].UniqId;
                        worksheet.Cell(rowNum, "B").Value = newQarsafari[r].LiterId;
                        worksheet.Cell(rowNum, "C").Value = newQarsafari[r].PhotoN;
                        worksheet.Cell(rowNum, "D").Value = newQarsafari[r].Region;
                        worksheet.Cell(rowNum, "E").Value = newQarsafari[r].Municipality;
                        worksheet.Cell(rowNum, "F").Value = newQarsafari[r].AdmMun;
                        worksheet.Cell(rowNum, "G").Value = newQarsafari[r].CityTownVillage;
                        worksheet.Cell(rowNum, "H").Value = newQarsafari[r].LandAreaSqM;
                        worksheet.Cell(rowNum, "I").Value = newQarsafari[r].LandAreaHa;
                        worksheet.Cell(rowNum, "J").Value = newQarsafari[r].Shrubbery;
                        worksheet.Cell(rowNum, "K").Value = newQarsafari[r].WoodyPlantPercent;
                        worksheet.Cell(rowNum, "L").Value = newQarsafari[r].WoodyPlantQuantity;
                        worksheet.Cell(rowNum, "M").Value = newQarsafari[r].VarjisFarti;
                        worksheet.Cell(rowNum, "N").Value = newQarsafari[r].WoodyPlantSpecies;

                        if (newQarsafari[r].WoodyPlantQuantity == 0)
                        {
                            worksheet.Cell(rowNum, "O").Value = 0;
                            worksheet.Cell(rowNum, "P").Value = 0;
                            worksheet.Cell(rowNum, "Q").Value = 0;
                        }
                        else
                        {
                            worksheet.Cell(rowNum, "O").Value = Math.Round(Convert.ToDouble(newQarsafari[r].InGoodCondition ?? 0), 1);
                            worksheet.Cell(rowNum, "P").Value = Math.Round(Convert.ToDouble(newQarsafari[r].ChoppedDown ?? 0), 1);
                            worksheet.Cell(rowNum, "Q").Value = Math.Round(Convert.ToDouble(newQarsafari[r].Rampike ?? 0), 1);
                        }
                        worksheet.Cell(rowNum, "R").Value = newQarsafari[r].SpeciesMediumAge;
                        worksheet.Cell(rowNum, "S").Value = newQarsafari[r].Note;
                        worksheet.Cell(rowNum, "T").Value = newQarsafari[r].Company;
                        worksheet.Cell(rowNum, "U").Value = newQarsafari[r].FieldOperator;
                        worksheet.Cell(rowNum, "V").Value = newQarsafari[r].Date;
                        worksheet.Cell(rowNum, "W").Value = newQarsafari[r].GisOperator;
                        worksheet.Cell(rowNum, "X").Value = newQarsafari[r].DaTe1;
                        worksheet.Cell(rowNum, "Y").Value = newQarsafari[r].OverlapCadCode;
                        worksheet.Cell(rowNum, "Z").Value = newQarsafari[r].Owner;
                        worksheet.Cell(rowNum, "AA").Value = newQarsafari[r].LegalPerson;
                        worksheet.Cell(rowNum, "AB").Value = newQarsafari[r].Owners;
                        worksheet.Cell(rowNum, "AC").Value = newQarsafari[r].LandFieldOperator;
                        worksheet.Cell(rowNum, "AD").Value = newQarsafari[r].Note1;
                        worksheet.Cell(rowNum, "AE").Value = newQarsafari[r].Date2;
                        worksheet.Cell(rowNum, "AF").Value = newQarsafari[r].LandGisOperator;
                        worksheet.Cell(rowNum, "AG").Value = newQarsafari[r].Note11;
                        worksheet.Cell(rowNum, "AH").Value = newQarsafari[r].Date3;
                        worksheet.Cell(rowNum, "AI").Value = newQarsafari[r].CadCod;
                        worksheet.Cell(rowNum, "AJ").Value = newQarsafari[r].UniqIdOld;
                        worksheet.Cell(rowNum, "AK").Value = newQarsafari[r].UniqIdNew;
                        worksheet.Cell(rowNum, "AL").Value = newQarsafari[r].Uid;
                        worksheet.Cell(rowNum, "AM").Value = newQarsafari[r].Id;
                        worksheet.Cell(rowNum, "AN").Value = qarsafariGrouped?.UniqId;
                    }

                    // Create result folder if it doesn't exist
                    string resultFolderPath = Path.Combine(ExcelDestinationPath, "result");
                    if (!Directory.Exists(resultFolderPath))
                    {
                        Directory.CreateDirectory(resultFolderPath);
                    }

                    workbook.SaveAs(Path.Combine(resultFolderPath, "rootExcel.xlsx"));
                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ექსელში ჩაწერისას!" + ex.Message
                };
            }

        }


        ///////აქ იქმნება result ფოლდერი თუ შექმნილი არაა და კოპირდება ძირი ექსელის ფაილი სადაც იწერება გადანომრილი uniqId ები 
        public Result<bool> copyOldExcelOriginal(ExcelReadDTO excelReadDTO)
        {
            Result<bool> result = new Result<bool>();

            try
            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
                var ExcelPath = excelReadDTO.ExcelPath;
                string excelDirectoryPath = Path.GetDirectoryName(ExcelPath);
                string fileName = Path.GetFileName(ExcelPath);

                string resultFolderPath = Path.Combine(excelDirectoryPath, "result");
                string destinationFilePath = Path.Combine(resultFolderPath, fileName);

                // Check if the result folder exists, if not, create it
                if (!Directory.Exists(resultFolderPath))
                {
                    Directory.CreateDirectory(resultFolderPath);
                }

                // Copy the file to the destination folder
                File.Copy(ExcelPath, destinationFilePath, true); // 'true' to overwrite if the file already exists

                // Copy Access file if provided
                if (!string.IsNullOrEmpty(excelReadDTO.AccessFilePath) && File.Exists(excelReadDTO.AccessFilePath))
                {
                    string accessFileName = Path.GetFileName(excelReadDTO.AccessFilePath);
                    string destinationAccessFilePath = Path.Combine(resultFolderPath, accessFileName);
                    File.Copy(excelReadDTO.AccessFilePath, destinationAccessFilePath, true);
                }

                // Copy images folder if provided
                if (!string.IsNullOrEmpty(excelReadDTO.FolderPath) && Directory.Exists(excelReadDTO.FolderPath))
                {
                    string folderName = Path.GetFileName(excelReadDTO.FolderPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                    if (string.IsNullOrEmpty(folderName))
                    {
                        folderName = "Images"; // Default name if folder path ends with separator
                    }
                    string destinationFolderPath = Path.Combine(resultFolderPath, folderName);
                    
                    // Copy directory recursively
                    CopyDirectory(excelReadDTO.FolderPath, destinationFolderPath, true);
                }

                // Open the copied Excel file and modify it using ClosedXML
                using (var workbook = new XLWorkbook(destinationFilePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // Find the column index for "Uniq_Id_New"
                    int column = 1;
                    var headerRow = worksheet.FirstRow();
                    while (column <= 100) // reasonable limit
                    {
                        var cellValue = headerRow.Cell(column).GetValue<string>();
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            // If "Uniq_Id_New" header does not exist, add it
                            headerRow.Cell(column).Value = "Uniq_Id_New";
                            break;
                        }
                        if (cellValue == "Uniq_Id_New")
                        {
                            break;
                        }
                        column++;
                    }

                    // Find the column index for "Uniq_Id_MDB"
                    int mdbColumn = column + 1;
                    while (mdbColumn <= 100)
                    {
                        var cellValue = headerRow.Cell(mdbColumn).GetValue<string>();
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            // If "Uniq_Id_MDB" header does not exist, add it
                            headerRow.Cell(mdbColumn).Value = "Uniq_Id_MDB";
                            break;
                        }
                        if (cellValue == "Uniq_Id_MDB")
                        {
                            break;
                        }
                        mdbColumn++;
                    }

                    // Read Excel data into a dictionary for faster lookup
                    Dictionary<string, string> excelData = new Dictionary<string, string>();
                    var lastRow = worksheet.LastRowUsed();
                    int rowCount = lastRow != null ? lastRow.RowNumber() : 0;

                    for (int i = 2; i <= rowCount; i++) // Assuming data starts from row 2
                    {
                        var row = worksheet.Row(i);
                        string uniqId = row.Cell(1).GetValue<string>();
                        string litterId = row.Cell(2).GetValue<string>();

                        // Only add to dictionary if both uniqId and litterId are not null
                        if (!string.IsNullOrEmpty(uniqId) && !string.IsNullOrEmpty(litterId))
                        {
                            excelData.Add($"{uniqId}_{litterId}", row.Cell(column).GetValue<string>() ?? "");
                        }
                    }

                    // Fetch data from database
                    var dbData = geographicDynamicDbContext.QarsafariGroupeds.ToList();

                    // Update Excel with fetched data
                    for (int i = 2; i <= rowCount; i++)
                    {
                        var row = worksheet.Row(i);
                        string excelUniqId = row.Cell(1).GetValue<string>();
                        string excelLitterId = row.Cell(2).GetValue<string>();

                        if (!string.IsNullOrEmpty(excelUniqId) && !string.IsNullOrEmpty(excelLitterId))
                        {
                            var matchedData = dbData.FirstOrDefault(item =>
                                item.UniqIdOld.ToString() == excelUniqId &&
                                item.LiterId.ToString() == excelLitterId);

                            if (matchedData != null)
                            {
                                // Write new uniqId from DB
                                row.Cell(column).Value = matchedData.UniqId;

                                var match = excelDataList.FirstOrDefault(m =>
                                            m.Uniq_ID_gadanomrili == matchedData.UniqId.ToString() &&
                                            m.Litter_Id == matchedData.LiterId.ToString());

                                if (match != null)
                                {
                                    row.Cell(mdbColumn).Value = match.Uniq_Id_MDB;
                                }
                                else
                                {
                                    // Optional logging in case MDB is missing
                                    Console.WriteLine($"No MDB match found for row {i}, UniqId: {matchedData.UniqId}");
                                }
                            }
                            else
                            {
                                // Handle rows with no database match
                                Console.WriteLine($"No DB match found for row {i}, UniqId: {excelUniqId}, LitterId: {excelLitterId}");
                            }
                        }
                    }

                    // Save changes
                    workbook.Save();
                }

                result.Success = true;
                result.StatusCode = System.Net.HttpStatusCode.OK;
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occurred during file operations or Excel manipulation
                result.Success = false;
                result.StatusCode = System.Net.HttpStatusCode.InternalServerError;
                result.Message = $"ძირი ექსელის, Access ფაილის და სურათების ფოლდერის გადაკოპირებისა და ექსელში ახალი UniqId ჩაწერისას მოხდა შეცდომა. შეცდომის რიგი: {ex.Message}";
                // Optionally log the exception details for troubleshooting
            }

            return result;
        }


        // Helper method to copy directory recursively
        private void CopyDirectory(string sourceDir, string destinationDir, bool overwrite)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDir);
            DirectoryInfo[] dirs = dir.GetDirectories();

            // Create the destination directory if it doesn't exist
            if (!Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
            }

            // Copy all files
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string tempPath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(tempPath, overwrite);
            }

            // Copy all subdirectories recursively
            foreach (DirectoryInfo subdir in dirs)
            {
                string tempPath = Path.Combine(destinationDir, subdir.Name);
                CopyDirectory(subdir.FullName, tempPath, overwrite);
            }
        }

        // ამრგვალებს 5 ის ჯერადზე გადაცემულ რიცხვს
        static int RoundToNearest(double number)
        {
            try
            {

                int i = Convert.ToInt32(number);
                return (i % 5) == 0 ? i : (i % 5) >= 2.5 ? i + 5 - (i % 5) : i - (i % 5);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"მოხდა შეცდომა რიცხვი გაიყო ნულზე საშუალო ხმოვანობაში სავარაუდოდ ! : {ex.Message}");
                return 0;
            }
        }
        #endregion
    }
}
