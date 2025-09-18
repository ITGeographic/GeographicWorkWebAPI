using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamicWebAPI.Wrappers;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace GeographicDynamic_DAL.Models.WindbreakMethods
{
    public class WindbreakMethods
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
            Application xlApp = new Application();
            try
            {
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelPath);
                //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1 ];
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
                //label2.Visible = true;
                Microsoft.Office.Interop.Excel.Range lastCell = xlWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int colCount = lastCell.Row;
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

                    // თუ უნიკიდ ან ლიტერ აიდი ცარიელია მაშინ ჩაიწერება false თუ არაა ცარიელი მაშინ true
                    if (String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, "A"].Value2)) || String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, "B"].Value2)))
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


                            object cellValue = xlRange.Cells[i, columnName.ColN].Value2;
                            if (cellValue != null)
                            {


                                Type cellType = cellValue.GetType();
                                PropertyInfo propertyInfo = typeof(Qarsafari).GetProperty(columnName.Sqlname);
                                if (propertyInfo != null)
                                {

                                    // Handle conversion based on cell type
                                    if (cellType == typeof(double))
                                    {
                                        propertyInfo.SetValue(qarsafari, (double)cellValue);
                                    }
                                    else if (cellType == typeof(string))
                                    {
                                        propertyInfo.SetValue(qarsafari, cellValue);
                                    }
                                    else if (cellType == typeof(DateTime))
                                    {
                                        DateTime dateTimeValue;
                                        if (DateTime.TryParse((string)cellValue, out dateTimeValue))
                                        {
                                            propertyInfo.SetValue(qarsafari, dateTimeValue);
                                        }
                                        // Handle DateTime conversion if necessary
                                    }
                                    // Add other type conversions as necessary
                                }

                            }
                            //else
                            //{
                            //    cellValue = (int)cellValue + 1;
                            //    Type cellType = cellValue.GetType();
                            //    PropertyInfo propertyInfo = typeof(Qarsafari).GetProperty(columnName.Sqlname);
                            //    if (propertyInfo != null)
                            //    {

                            //        // Handle conversion based on cell type
                            //        if (cellType == typeof(double))
                            //        {
                            //            propertyInfo.SetValue(qarsafari, (double)cellValue);
                            //        }
                            //        else if (cellType == typeof(string))
                            //        {
                            //            propertyInfo.SetValue(qarsafari, cellValue);
                            //        }
                            //        else if (cellType == typeof(DateTime))
                            //        {
                            //            DateTime dateTimeValue;
                            //            if (DateTime.TryParse((string)cellValue, out dateTimeValue))
                            //            {
                            //                propertyInfo.SetValue(qarsafari, dateTimeValue);
                            //            }
                            //            // Handle DateTime conversion if necessary
                            //        }
                            //        // Add other type conversions as necessary
                            //    }
                            //}
                        }
                    }
                    //qarsafari.UniqId = Convert.ToDouble(xlRange.Cells[i, "A"].Value2);
                    //qarsafari.LiterId = Convert.ToDouble(xlRange.Cells[i, "B"].Value2);
                    //qarsafari.PhotoN = xlRange.Cells[i, "C"].Value2;
                    //qarsafari.Region = xlRange.Cells[i, "D"].value2;
                    //qarsafari.Municipality = xlRange.Cells[i, "E"].Value2;
                    //qarsafari.Shrubbery = Convert.ToDouble(xlRange.Cells[i, "J"].Value2);
                    ////qarsafari.WoodyPlantPercent = Convert.ToDouble(xlRange.Cells[i,"K"].Value2);
                    //qarsafari.WoodyPlantQuantity = Convert.ToDouble(xlRange.Cells[i, "L"].Value2);
                    //qarsafari.WoodyPlantSpecies = xlRange.Cells[i, "M"].Value2;
                    //qarsafari.InGoodCondition = Convert.ToDouble(xlRange.Cells[i, "N"].Value2);
                    //qarsafari.ChoppedDown = Convert.ToDouble(xlRange.Cells[i, "O"].Value2);
                    //qarsafari.Rampike = Convert.ToDouble(xlRange.Cells[i, "P"].Value2);
                    //qarsafari.SpeciesMediumAge = Convert.ToDouble(xlRange.Cells[i, "Q"].Value2);
                    //qarsafari.Company = xlRange.Cells[i, "S"].Value2;
                    //qarsafari.FieldOperator = xlRange.Cells[i, "T"].Value2;
                    //qarsafari.Owners = xlRange.Cells[i, "AA"].Value2;
                    //qarsafari.LandFieldOperator = xlRange.Cells[i, "AB"].Value2;
                    //qarsafari.Note1 = xlRange.Cells[i, "AC"].Value2;
                    //qarsafari.Date2 = xlRange.Cells[i, "AD"].Value2;
                    //qarsafari.LandGisOperator = xlRange.Cells[i, "AE"].Value2;
                    //qarsafari.Note11 = xlRange.Cells[i, "AF"].Value2;
                    //qarsafari.Date3 = xlRange.Cells[i, "AG"].Value2;
                    //qarsafari.CadCod = xlRange.Cells[i, "AH"].Value2;

                    GeographicDynamicDbContext.Qarsafaris.Add(qarsafari); // ახალი ობიექტის დამატება ბაზაში
                    GeographicDynamicDbContext.SaveChanges(); // ცვლილებების შენახვა


                    //label2.Text = Convert.ToString($"{i - 1} / {colCount - 1}");
                    //GeographicDynamicDbContext.PerformStep();

                }
                xlApp.Application.Quit();
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                xlApp.Application.Quit();
                return new Result<bool> { Success = false, StatusCode = System.Net.HttpStatusCode.BadGateway, Message = "შეცდომა მოხდა " + ex.Message };
                throw;
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
                        System.Data.DataRow[] rows = dataTable.Select($"{uniqid} = {item.UniqIdOld} AND {literid} = {item.LiterId}");
                        if (rows.Length > 1)
                        {
                            var test = "test";
                        }
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
                                //////////row["UNIQ_ID"] = item.UniqId;
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

                // Second: reopen connection for shifting columns
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    // Step 1: check for NULLs in Uniq_ID_NEW_Gadanomrili
                    string nullCheckQuery = $@"SELECT COUNT(*) FROM [{AccessSheetName}] WHERE Uniq_ID_NEW_Gadanomrili IS NULL";
                    using (OleDbCommand nullCheckCmd = new OleDbCommand(nullCheckQuery, connection))
                    {
                        int nullCount = (int)nullCheckCmd.ExecuteScalar();
                        if (nullCount > 0)
                        {
                            throw new Exception($"Found {nullCount} rows where Uniq_ID_NEW_Gadanomrili is NULL. Update aborted.");
                        }
                    }

                                    // Step 2: shift columns if safe
                                    string updateShiftQuery = $@"
                        UPDATE [{AccessSheetName}]
                        SET 
                            UNIQ_ID = Uniq_ID_NEW_Gadanomrili
                    ";
                    using (OleDbCommand updateShiftCmd = new OleDbCommand(updateShiftQuery, connection))
                    {
                        updateShiftCmd.ExecuteNonQuery();
                    }

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

            //List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).ToList();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook = null;
            ExcelApp.Visible = false;
            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ExcelWorkSheet = ExcelWorkBook.Worksheets[1] as Worksheet;




            try
            {



                // ამით ივსება სათაურების ველები 
                ExcelWorkSheet.Cells[1, "A"] = "UNIQ_ID";
                ExcelWorkSheet.Cells[1, "B"] = "Liter_ID";
                ExcelWorkSheet.Cells[1, "C"] = "Photo_N";
                ExcelWorkSheet.Cells[1, "C"].EntireColumn.NumberFormat = "@";
                ExcelWorkSheet.Cells[1, "D"] = "Region";
                ExcelWorkSheet.Cells[1, "E"] = "Municipality";
                ExcelWorkSheet.Cells[1, "F"] = "Adm_Mun";
                ExcelWorkSheet.Cells[1, "G"] = "City_Town_Village";
                ExcelWorkSheet.Cells[1, "H"] = "Land_Area_Sq_M";
                ExcelWorkSheet.Cells[1, "I"] = "Land_Area_Ha";
                ExcelWorkSheet.Cells[1, "J"] = "Shrubbery";
                ExcelWorkSheet.Cells[1, "K"] = "Woody_Plant_Percent";
                ExcelWorkSheet.Cells[1, "L"] = "Woody_Plant_Quantity";
                ExcelWorkSheet.Cells[1, "M"] = "Woody_Plant_Spices";
                ExcelWorkSheet.Cells[1, "N"] = "VarjisFarti";
                ExcelWorkSheet.Cells[1, "O"] = "In_Good_Condition";
                ExcelWorkSheet.Cells[1, "P"] = "Chopped_down";
                ExcelWorkSheet.Cells[1, "Q"] = "Rampike";
                ExcelWorkSheet.Cells[1, "R"] = "Spices_Medium_Age";
                ExcelWorkSheet.Cells[1, "S"] = "Note_";
                ExcelWorkSheet.Cells[1, "T"] = "Company";
                ExcelWorkSheet.Cells[1, "U"] = "Field_Operator";
                ExcelWorkSheet.Cells[1, "V"] = "Date_";
                ExcelWorkSheet.Cells[1, "V"].EntireColumn.NumberFormat = "@"; //ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "W"] = "Gis_Operator";
                ExcelWorkSheet.Cells[1, "X"] = "DaTe_1";
                ExcelWorkSheet.Cells[1, "X"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "Y"] = "Overlap_CAD_CODE";
                ExcelWorkSheet.Cells[1, "Z"] = "Owner";
                ExcelWorkSheet.Cells[1, "AA"] = "Legal_person";
                ExcelWorkSheet.Cells[1, "AB"] = "Owners";
                ExcelWorkSheet.Cells[1, "AC"] = "Land_Field_Operator";
                ExcelWorkSheet.Cells[1, "AD"] = "Note1";
                ExcelWorkSheet.Cells[1, "AE"] = "Date_2";
                ExcelWorkSheet.Cells[1, "AF"] = "Land_Gis_Operator";
                ExcelWorkSheet.Cells[1, "AG"] = "Note1_1";
                ExcelWorkSheet.Cells[1, "AH"] = "Date_3";
                ExcelWorkSheet.Cells[1, "AI"] = "CAD_COD";
                ExcelWorkSheet.Cells[1, "AJ"] = "UNIQ_ID_OLD";
                ExcelWorkSheet.Cells[1, "AK"] = "UNIQ_ID_NEW";
                ExcelWorkSheet.Cells[1, "AL"] = "UID";
                ExcelWorkSheet.Cells[1, "AM"] = "ID";

                // ამ ციკლით ივსება Rows სათაურების ქვეშ 
                for (int r = 0; r < qarsafaris.Count(); r++) //r stands for ExcelRow and c for ExcelColumn
                {


                    QarsafariGrouped qarsafariGrouped = GeographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.UniqId == qarsafaris[r].UniqId);
                    //ExcelWorkSheet.Cells[r + 2, "A"] = qarsafaris[r].UniqId;
                    if (qarsafaris[r].IsUniqLiterNull == "true")
                    {
                        ExcelWorkSheet.Cells[r + 2, "A"] = qarsafaris[r].UniqId;
                        ExcelWorkSheet.Cells[r + 2, "B"] = qarsafaris[r].LiterId;
                        ExcelWorkSheet.Cells[r + 2, "AJ"] = qarsafaris[r].UniqIdOld;
                        ExcelWorkSheet.Cells[r + 2, "K"] = Math.Round(Convert.ToDouble(qarsafariGrouped.WoodyPlantPercent), 1);
                        ExcelWorkSheet.Cells[r + 2, "C"] = qarsafariGrouped.PhotoN; // ფოტოები მოდის დაგრუპულიდან 
                        ExcelWorkSheet.Cells[r + 2, "D"] = qarsafaris[r].Region;
                        ExcelWorkSheet.Cells[r + 2, "E"] = qarsafaris[r].Municipality;
                        ExcelWorkSheet.Cells[r + 2, "F"] = qarsafaris[r].AdmMun;
                        ExcelWorkSheet.Cells[r + 2, "G"] = qarsafaris[r].CityTownVillage;
                        ExcelWorkSheet.Cells[r + 2, "H"] = qarsafaris[r].LandAreaSqM;
                        ExcelWorkSheet.Cells[r + 2, "I"] = qarsafaris[r].LandAreaHa;
                        ExcelWorkSheet.Cells[r + 2, "V"] = qarsafariGrouped.Date;
                        ExcelWorkSheet.Cells[r + 2, "Z"] = qarsafaris[r].Owner;
                    }




                    ExcelWorkSheet.Cells[r + 2, "J"] = qarsafaris[r].Shrubbery;
                    ExcelWorkSheet.Cells[r + 2, "L"] = qarsafaris[r].WoodyPlantQuantity;
                    ExcelWorkSheet.Cells[r + 2, "M"] = qarsafaris[r].WoodyPlantSpecies;
                    ExcelWorkSheet.Cells[r + 2, "N"] = qarsafaris[r].VarjisFarti;

                    // მოწმდება თუ სადმე sumofGoodChoppedRampike განსხვავდება 0-ს ან 100-ს იდეაში 0.1 ან მეტია ან ნაკლები და მაგის მიხედვით 
                    // ხორციელდება გამოკლება ან მიმატება 0.1-ის 
                    var sumofGoodChoppedRampike = Math.Round(Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1)
                        + Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1)
                        + Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1), 1);

                    switch (sumofGoodChoppedRampike)
                    {
                        case 99.9:
                            ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1) + 0.1;
                            ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                            ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            break;
                        case 100.1:
                            if (qarsafaris[r].InGoodCondition > 0)
                            {
                                ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1) - 0.1;
                                ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                                ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            }
                            else if (qarsafaris[r].ChoppedDown > 0)
                            {
                                ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                                ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1) - 0.1;
                                ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            }
                            else
                            {
                                ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                                ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                                ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1) - 0.1;
                            }
                            break;
                        default:
                            ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                            ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                            ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            break;
                    }
                    //ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                    //ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                    //ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                    ExcelWorkSheet.Cells[r + 2, "R"] = qarsafaris[r].SpeciesMediumAge;
                    ExcelWorkSheet.Cells[r + 2, "S"] = qarsafaris[r].Note;
                    ExcelWorkSheet.Cells[r + 2, "T"] = qarsafaris[r].Company;
                    ExcelWorkSheet.Cells[r + 2, "U"] = qarsafaris[r].FieldOperator;
                    //ExcelWorkSheet.Cells[r + 2, "V"] = qarsafaris[r].Date;
                    ExcelWorkSheet.Cells[r + 2, "W"] = qarsafaris[r].GisOperator;
                    ExcelWorkSheet.Cells[r + 2, "X"] = qarsafaris[r].DaTe1;
                    ExcelWorkSheet.Cells[r + 2, "Y"] = qarsafaris[r].OverlapCadCode;

                    ExcelWorkSheet.Cells[r + 2, "AA"] = qarsafaris[r].LegalPerson;
                    ExcelWorkSheet.Cells[r + 2, "AB"] = qarsafaris[r].Owners;
                    ExcelWorkSheet.Cells[r + 2, "AC"] = qarsafaris[r].LandFieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "AD"] = qarsafaris[r].Note1;
                    ExcelWorkSheet.Cells[r + 2, "AE"] = qarsafaris[r].Date2;
                    ExcelWorkSheet.Cells[r + 2, "AF"] = qarsafaris[r].LandGisOperator;
                    ExcelWorkSheet.Cells[r + 2, "AG"] = qarsafaris[r].Note11;
                    ExcelWorkSheet.Cells[r + 2, "AH"] = qarsafaris[r].Date3;
                    ExcelWorkSheet.Cells[r + 2, "AI"] = qarsafaris[r].CadCod;
                    ExcelWorkSheet.Cells[r + 2, "AK"] = qarsafaris[r].UniqIdNew;
                    ExcelWorkSheet.Cells[r + 2, "AL"] = qarsafaris[r].Uid;
                    ExcelWorkSheet.Cells[r + 2, "AM"] = qarsafaris[r].Id;
                }

                ExcelWorkBook.Worksheets[1].Name = "დიდი-ექსელი";//Renaming the Sheet1 to MySheet
                ExcelWorkBook.SaveAs(ExcelDestinationPath + $"\\{ExcelName}.xlsx");
                ExcelWorkBook.Close();
                ExcelApp.Quit();
                Marshal.ReleaseComObject(ExcelWorkSheet);
                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }

            catch (Exception ex)
            {
                ExcelWorkBook.Close();
                ExcelApp.Quit();
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
        public Result<bool> WriteToExcelGrouped(List<QarsafariGrouped> qarsafariGroupeds, string ExcelDestinationPath, string ExcelName,string AccessPath,string AccessSheetName)
        {

            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext(); //უკავშირდება კონტექსტს რომ გაიგოს ცხრილები SQL-დან 

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application(); //იქმნება აპლიკაცია წინასწარ 
            Workbook ExcelWorkBook = null; // წინასწარ იქმნება ვორკბუკის ცვლადი რომ შემდეგში გამოვიყენოთ 
            Worksheet ExcelWorkSheet = null; // ასევე წინასწარ იქმნება შიტის ცვლადი რომ გამოვიყენოთ შემდეგ 
            ExcelApp.Visible = false; // აქ ვანიჭებთ ექსელის ფანჯარას რომ გამოჩნდეს პროგრამის მსვლელობა 
            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);  // იქმნება ახალი ექსელის workbook რომელშიც გვაქ ერთი შიტი და ამ შიტს ვიყენებთ სამომავლოდ 


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
               





                    ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; // აქ ვირჩევთ სამუშაო შიტს ვორკბუკიდან(ექსელიდან) ჩვენ შემთხვევაში ერთია და მაგიტომ გვაქ Worksheets[1} ინდექსად 1 


                // ამ კოდის ფრაგმენტებში ივსება სათაურის ველები სხვა სიტყვებით რომ ვთქვათ პირველ row-ში იწერება მნიშვნელობები რამდენი სვეტიც გვაქ (column) 
                ExcelWorkSheet.Cells[1, "A"] = "UNIQ_ID";
                ExcelWorkSheet.Cells[1, "B"] = "Liter_ID";
                ExcelWorkSheet.Cells[1, "C"] = "Photo_N";
                ExcelWorkSheet.Cells[1, "C"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "D"] = "REGION";
                ExcelWorkSheet.Cells[1, "E"] = "Municipality";
                ExcelWorkSheet.Cells[1, "F"] = "Adm_Mun";
                ExcelWorkSheet.Cells[1, "G"] = "City_Town_Village";
                ExcelWorkSheet.Cells[1, "H"] = "Land_Area_Sq_m";
                ExcelWorkSheet.Cells[1, "I"] = "Land_Area_Ha";
                ExcelWorkSheet.Cells[1, "J"] = "shrubbery";
                ExcelWorkSheet.Cells[1, "K"] = "Woody_plant_percent";
                ExcelWorkSheet.Cells[1, "L"] = "Woody_plant_quantity";
                ExcelWorkSheet.Cells[1, "M"] = "VarjisFarti";
                ExcelWorkSheet.Cells[1, "N"] = "woody_plant_species";
                ExcelWorkSheet.Cells[1, "O"] = "In_good_condition";
                ExcelWorkSheet.Cells[1, "P"] = "chopped_down";
                ExcelWorkSheet.Cells[1, "Q"] = "rampike";
                ExcelWorkSheet.Cells[1, "R"] = "species_medium_age";
                ExcelWorkSheet.Cells[1, "S"] = "Note_";
                ExcelWorkSheet.Cells[1, "T"] = "Company";
                ExcelWorkSheet.Cells[1, "U"] = "Field_Operator";
                ExcelWorkSheet.Cells[1, "V"] = "Date_";
                ExcelWorkSheet.Cells[1, "V"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "W"] = "Gis_Operator";
                ExcelWorkSheet.Cells[1, "X"] = "DaTe_1";
                ExcelWorkSheet.Cells[1, "X"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "Y"] = "Overlap_CAD_CODE";
                ExcelWorkSheet.Cells[1, "Z"] = "Owner";
                ExcelWorkSheet.Cells[1, "AA"] = "Legal_person";
                ExcelWorkSheet.Cells[1, "AB"] = "Owners";
                ExcelWorkSheet.Cells[1, "AC"] = "Land_Field_Operator";
                ExcelWorkSheet.Cells[1, "AD"] = "Note1";
                ExcelWorkSheet.Cells[1, "AE"] = "Date_2";
                ExcelWorkSheet.Cells[1, "AF"] = "Land_Gis_Operator";
                ExcelWorkSheet.Cells[1, "AG"] = "Note1_1";
                ExcelWorkSheet.Cells[1, "AH"] = "Date_3";
                ExcelWorkSheet.Cells[1, "AI"] = "CAD_COD";
                ExcelWorkSheet.Cells[1, "AJ"] = "UNIQ_ID_OLD";
                ExcelWorkSheet.Cells[1, "AK"] = "UNIQ_ID_NEW";
                ExcelWorkSheet.Cells[1, "AL"] = "UID";
                ExcelWorkSheet.Cells[1, "AM"] = "ID";
                ExcelWorkSheet.Cells[1, "AN"] = "Uniq_ID_MDB";

                for (var r = 0; r < qarsafariGroupeds.Count(); r++) // კეთდება ციკლი იმისთვის რო დაიაროს სათითაო ველი და ჩაიწეროს ექსელის შიტში 
                                                                    // R ამ შემთხვევაში ნიშნავს RowNumbers რომ ჩაწერა დაიწყოს მეროე რიგიდან რადგან პირველიში სვეტების სახელები წერია 
                                                                    // ყოველ იტერაციაზე R-ს ერთი ემატება რის გამოც შემდეგ რიგში გადადის ინფორმაციის შევსება 
                {
                    // ციკლის შიგნით იწერება რომელ სვეტში რა ინფორმაცია ჩაიწეროს 

                    ExcelWorkSheet.Cells[r + 2, "A"] = qarsafariGroupeds[r].UniqId;
                    ExcelWorkSheet.Cells[r + 2, "B"] = qarsafariGroupeds[r].LiterId;
                    ExcelWorkSheet.Cells[r + 2, "C"] = qarsafariGroupeds[r].PhotoN;
                    ExcelWorkSheet.Cells[r + 2, "D"] = qarsafariGroupeds[r].Region;
                    ExcelWorkSheet.Cells[r + 2, "E"] = qarsafariGroupeds[r].Municipality;
                    ExcelWorkSheet.Cells[r + 2, "F"] = qarsafariGroupeds[r].AdmMun;
                    ExcelWorkSheet.Cells[r + 2, "G"] = qarsafariGroupeds[r].CityTownVillage;
                    ExcelWorkSheet.Cells[r + 2, "H"] = qarsafariGroupeds[r].LandAreaSqM;
                    ExcelWorkSheet.Cells[r + 2, "I"] = qarsafariGroupeds[r].LandAreaHa;
                    ExcelWorkSheet.Cells[r + 2, "J"] = qarsafariGroupeds[r].Shrubbery;
                    ExcelWorkSheet.Cells[r + 2, "K"] = qarsafariGroupeds[r].WoodyPlantPercent;
                    ExcelWorkSheet.Cells[r + 2, "L"] = qarsafariGroupeds[r].WoodyPlantQuantity;
                    ExcelWorkSheet.Cells[r + 2, "M"] = qarsafariGroupeds[r].VarjisFarti;
                    ExcelWorkSheet.Cells[r + 2, "N"] = qarsafariGroupeds[r].WoodyPlantSpecies;

                    // სადაც სახეობა არ გვიწერია და ხეხილის რაოდენობა იქ იწერება კარგ მდომარეობაში 0 
                    if (qarsafariGroupeds[r].WoodyPlantQuantity == 0)
                    {
                        ExcelWorkSheet.Cells[r + 2, "O"] = 0;
                        ExcelWorkSheet.Cells[r + 2, "P"] = 0;
                        ExcelWorkSheet.Cells[r + 2, "Q"] = 0;
                    }
                    else
                    {
                        ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].InGoodCondition), 1);
                        ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].ChoppedDown), 1);
                        ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].Rampike), 1);
                    }
                    ExcelWorkSheet.Cells[r + 2, "R"] = qarsafariGroupeds[r].SpeciesMediumAge;
                    ExcelWorkSheet.Cells[r + 2, "S"] = qarsafariGroupeds[r].Note;
                    ExcelWorkSheet.Cells[r + 2, "T"] = qarsafariGroupeds[r].Company;
                    ExcelWorkSheet.Cells[r + 2, "U"] = qarsafariGroupeds[r].FieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "V"] = qarsafariGroupeds[r].Date;
                    ExcelWorkSheet.Cells[r + 2, "W"] = qarsafariGroupeds[r].GisOperator;
                    ExcelWorkSheet.Cells[r + 2, "X"] = qarsafariGroupeds[r].DaTe1;
                    ExcelWorkSheet.Cells[r + 2, "Y"] = qarsafariGroupeds[r].OverlapCadCode;
                    ExcelWorkSheet.Cells[r + 2, "Z"] = qarsafariGroupeds[r].Owner;
                    ExcelWorkSheet.Cells[r + 2, "AA"] = qarsafariGroupeds[r].LegalPerson;
                    ExcelWorkSheet.Cells[r + 2, "AB"] = qarsafariGroupeds[r].Owners;
                    ExcelWorkSheet.Cells[r + 2, "AC"] = qarsafariGroupeds[r].LandFieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "AD"] = qarsafariGroupeds[r].Note1;
                    ExcelWorkSheet.Cells[r + 2, "AE"] = qarsafariGroupeds[r].Date2;
                    ExcelWorkSheet.Cells[r + 2, "AF"] = qarsafariGroupeds[r].LandGisOperator;
                    ExcelWorkSheet.Cells[r + 2, "AG"] = qarsafariGroupeds[r].Note11;
                    ExcelWorkSheet.Cells[r + 2, "AH"] = qarsafariGroupeds[r].Date3;
                    ExcelWorkSheet.Cells[r + 2, "AI"] = qarsafariGroupeds[r].CadCod;
                    ExcelWorkSheet.Cells[r + 2, "AJ"] = qarsafariGroupeds[r].UniqIdOld;
                    ExcelWorkSheet.Cells[r + 2, "AK"] = qarsafariGroupeds[r].UniqIdNew;
                    ExcelWorkSheet.Cells[r + 2, "AL"] = qarsafariGroupeds[r].Uid;
                    ExcelWorkSheet.Cells[r + 2, "AM"] = qarsafariGroupeds[r].Id;

                    string uniqIdGadanomrili = qarsafariGroupeds[r].UniqId.ToString();
                    if (accessData.ContainsKey(uniqIdGadanomrili))
                    {
                        ExcelWorkSheet.Cells[r + 2, "AN"] = accessData[uniqIdGadanomrili];
                    }
                }
                ExcelWorkBook.Worksheets[1].Name = "პატარა-ექსელი"; // ვარქმევთ ჩვენ მიერ ზევით შექმნილ შიტს სახელს 
                ExcelWorkBook.SaveAs(ExcelDestinationPath + $"\\{ExcelName}.xlsx"); // ვაძლევთ სახელს ექსელის ფაილს ჩვენ მიერ გადმოწოდებული ცვლადის მიხედვით 
                ExcelWorkBook.Close(); // იხურება ექსელის ფაილი ვეღარ მოვახდენთ მასზე ცვლილებას 
                ExcelApp.Quit(); //გამოვდივართ აპლიკაციიდან 

                Marshal.ReleaseComObject(ExcelWorkSheet);
                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }
            catch (Exception ex)
            {
                ExcelWorkBook.Close(); // იხურება ექსელის ფაილი ვეღარ მოვახდენთ მასზე ცვლილებას 
                ExcelApp.Quit(); //გამოვდივართ აპლიკაციიდან 
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

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application(); //იქმნება აპლიკაცია წინასწარ 
            Workbook ExcelWorkBook = null; // წინასწარ იქმნება ვორკბუკის ცვლადი რომ შემდეგში გამოვიყენოთ 
            Worksheet ExcelWorkSheet = null; // ასევე წინასწარ იქმნება შიტის ცვლადი რომ გამოვიყენოთ შემდეგ 
            ExcelApp.Visible = false; // აქ ვანიჭებთ ექსელის ფანჯარას რომ გამოჩნდეს პროგრამის მსვლელობა 
            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);  // იქმნება ახალი ექსელის workbook რომელშიც გვაქ ერთი შიტი და ამ შიტს ვიყენებთ სამომავლოდ 

            try
            {
                ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; // აქ ვირჩევთ სამუშაო შიტს ვორკბუკიდან(ექსელიდან) ჩვენ შემთხვევაში ერთია და მაგიტომ გვაქ Worksheets[1} ინდექსად 1 


                // ამ კოდის ფრაგმენტებში ივსება სათაურის ველები სხვა სიტყვებით რომ ვთქვათ პირველ row-ში იწერება მნიშვნელობები რამდენი სვეტიც გვაქ (column) 
                ExcelWorkSheet.Cells[1, "A"] = "UNIQ_ID";
                ExcelWorkSheet.Cells[1, "B"] = "Liter_ID";
                ExcelWorkSheet.Cells[1, "C"] = "Photo_N";
                //ExcelWorkSheet.Cells[1, "C"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "D"] = "REGION";
                ExcelWorkSheet.Cells[1, "E"] = "Municipality";
                ExcelWorkSheet.Cells[1, "F"] = "Adm_Mun";
                ExcelWorkSheet.Cells[1, "G"] = "City_Town_Village";
                ExcelWorkSheet.Cells[1, "H"] = "Land_Area_Sq_m";
                ExcelWorkSheet.Cells[1, "I"] = "Land_Area_Ha";
                ExcelWorkSheet.Cells[1, "J"] = "shrubbery";
                ExcelWorkSheet.Cells[1, "K"] = "Woody_plant_percent";
                ExcelWorkSheet.Cells[1, "L"] = "Woody_plant_quantity";
                ExcelWorkSheet.Cells[1, "M"] = "VarjisFarti";
                ExcelWorkSheet.Cells[1, "N"] = "woody_plant_species";
                ExcelWorkSheet.Cells[1, "O"] = "In_good_condition";
                ExcelWorkSheet.Cells[1, "P"] = "chopped_down";
                ExcelWorkSheet.Cells[1, "Q"] = "rampike";
                ExcelWorkSheet.Cells[1, "R"] = "species_medium_age";
                ExcelWorkSheet.Cells[1, "S"] = "Note_";
                ExcelWorkSheet.Cells[1, "T"] = "Company";
                ExcelWorkSheet.Cells[1, "U"] = "Field_Operator";
                ExcelWorkSheet.Cells[1, "V"] = "Date_";
                ExcelWorkSheet.Cells[1, "V"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "W"] = "Gis_Operator";
                ExcelWorkSheet.Cells[1, "X"] = "DaTe_1";
                ExcelWorkSheet.Cells[1, "X"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "Y"] = "Overlap_CAD_CODE";
                ExcelWorkSheet.Cells[1, "Z"] = "Owner";
                ExcelWorkSheet.Cells[1, "AA"] = "Legal_person";
                ExcelWorkSheet.Cells[1, "AB"] = "Owners";
                ExcelWorkSheet.Cells[1, "AC"] = "Land_Field_Operator";
                ExcelWorkSheet.Cells[1, "AD"] = "Note1";
                ExcelWorkSheet.Cells[1, "AE"] = "Date_2";
                ExcelWorkSheet.Cells[1, "AF"] = "Land_Gis_Operator";
                ExcelWorkSheet.Cells[1, "AG"] = "Note1_1";
                ExcelWorkSheet.Cells[1, "AH"] = "Date_3";
                ExcelWorkSheet.Cells[1, "AI"] = "CAD_COD";
                ExcelWorkSheet.Cells[1, "AJ"] = "UNIQ_ID_OLD";
                ExcelWorkSheet.Cells[1, "AK"] = "UNIQ_ID_NEW";
                ExcelWorkSheet.Cells[1, "AL"] = "UID";
                ExcelWorkSheet.Cells[1, "AM"] = "ID";
                ExcelWorkSheet.Cells[1, "AN"] = "Uniq_Id_NEW";

                for (var r = 0; r < newQarsafari.Count(); r++) // კეთდება ციკლი იმისთვის რო დაიაროს სათითაო ველი და ჩაიწეროს ექსელის შიტში 
                                                               // R ამ შემთხვევაში ნიშნავს RowNumbers რომ ჩაწერა დაიწყოს მეროე რიგიდან რადგან პირველიში სვეტების სახელები წერია 
                                                               // ყოველ იტერაციაზე R-ს ერთი ემატება რის გამოც შემდეგ რიგში გადადის ინფორმაციის შევსება 
                {

                    QarsafariGrouped qarsafariGrouped = geographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.UniqIdOld == newQarsafari[r].UniqId);
                    // ციკლის შიგნით იწერება რომელ სვეტში რა ინფორმაცია ჩაიწეროს 

                    ExcelWorkSheet.Cells[r + 2, "A"] = newQarsafari[r].UniqId;
                    ExcelWorkSheet.Cells[r + 2, "B"] = newQarsafari[r].LiterId;
                    ExcelWorkSheet.Cells[r + 2, "C"] = newQarsafari[r].PhotoN;
                    ExcelWorkSheet.Cells[r + 2, "D"] = newQarsafari[r].Region;
                    ExcelWorkSheet.Cells[r + 2, "E"] = newQarsafari[r].Municipality;
                    ExcelWorkSheet.Cells[r + 2, "F"] = newQarsafari[r].AdmMun;
                    ExcelWorkSheet.Cells[r + 2, "G"] = newQarsafari[r].CityTownVillage;
                    ExcelWorkSheet.Cells[r + 2, "H"] = newQarsafari[r].LandAreaSqM;
                    ExcelWorkSheet.Cells[r + 2, "I"] = newQarsafari[r].LandAreaHa;
                    ExcelWorkSheet.Cells[r + 2, "J"] = newQarsafari[r].Shrubbery;
                    ExcelWorkSheet.Cells[r + 2, "K"] = newQarsafari[r].WoodyPlantPercent;
                    ExcelWorkSheet.Cells[r + 2, "L"] = newQarsafari[r].WoodyPlantQuantity;
                    ExcelWorkSheet.Cells[r + 2, "M"] = newQarsafari[r].VarjisFarti;
                    ExcelWorkSheet.Cells[r + 2, "N"] = newQarsafari[r].WoodyPlantSpecies;

                    // სადაც სახეობა არ გვიწერია და ხეხილის რაოდენობა იქ იწერება კარგ მდომარეობაში 0 
                    if (newQarsafari[r].WoodyPlantQuantity == 0)
                    {
                        ExcelWorkSheet.Cells[r + 2, "O"] = 0;
                        ExcelWorkSheet.Cells[r + 2, "P"] = 0;
                        ExcelWorkSheet.Cells[r + 2, "Q"] = 0;
                    }
                    else
                    {
                        ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(newQarsafari[r].InGoodCondition), 1);
                        ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(newQarsafari[r].ChoppedDown), 1);
                        ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(newQarsafari[r].Rampike), 1);
                    }
                    ExcelWorkSheet.Cells[r + 2, "R"] = newQarsafari[r].SpeciesMediumAge;
                    ExcelWorkSheet.Cells[r + 2, "S"] = newQarsafari[r].Note;
                    ExcelWorkSheet.Cells[r + 2, "T"] = newQarsafari[r].Company;
                    ExcelWorkSheet.Cells[r + 2, "U"] = newQarsafari[r].FieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "V"] = newQarsafari[r].Date;
                    ExcelWorkSheet.Cells[r + 2, "W"] = newQarsafari[r].GisOperator;
                    ExcelWorkSheet.Cells[r + 2, "X"] = newQarsafari[r].DaTe1;
                    ExcelWorkSheet.Cells[r + 2, "Y"] = newQarsafari[r].OverlapCadCode;
                    ExcelWorkSheet.Cells[r + 2, "Z"] = newQarsafari[r].Owner;
                    ExcelWorkSheet.Cells[r + 2, "AA"] = newQarsafari[r].LegalPerson;
                    ExcelWorkSheet.Cells[r + 2, "AB"] = newQarsafari[r].Owners;
                    ExcelWorkSheet.Cells[r + 2, "AC"] = newQarsafari[r].LandFieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "AD"] = newQarsafari[r].Note1;
                    ExcelWorkSheet.Cells[r + 2, "AE"] = newQarsafari[r].Date2;
                    ExcelWorkSheet.Cells[r + 2, "AF"] = newQarsafari[r].LandGisOperator;
                    ExcelWorkSheet.Cells[r + 2, "AG"] = newQarsafari[r].Note11;
                    ExcelWorkSheet.Cells[r + 2, "AH"] = newQarsafari[r].Date3;
                    ExcelWorkSheet.Cells[r + 2, "AI"] = newQarsafari[r].CadCod;
                    ExcelWorkSheet.Cells[r + 2, "AJ"] = newQarsafari[r].UniqIdOld;
                    ExcelWorkSheet.Cells[r + 2, "AK"] = newQarsafari[r].UniqIdNew;
                    ExcelWorkSheet.Cells[r + 2, "AL"] = newQarsafari[r].Uid;
                    ExcelWorkSheet.Cells[r + 2, "AM"] = newQarsafari[r].Id;
                    ExcelWorkSheet.Cells[r + 2, "AN"] = qarsafariGrouped.UniqId;
                }
                ExcelWorkBook.Worksheets[1].Name = "პატარა-ექსელი"; // ვარქმევთ ჩვენ მიერ ზევით შექმნილ შიტს სახელს 
                ExcelWorkBook.SaveAs(ExcelDestinationPath + "rootExcel.xlsx"); // ვაძლევთ სახელს ექსელის ფაილს ჩვენ მიერ გადმოწოდებული ცვლადის მიხედვით 
                ExcelWorkBook.Close(); // იხურება ექსელის ფაილი ვეღარ მოვახდენთ მასზე ცვლილებას 
                ExcelApp.Quit(); //გამოვდივართ აპლიკაციიდან 

                Marshal.ReleaseComObject(ExcelWorkSheet);
                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }
            catch (Exception ex)
            {
                ExcelWorkBook.Close(); // იხურება ექსელის ფაილი ვეღარ მოვახდენთ მასზე ცვლილებას 
                ExcelApp.Quit(); //გამოვდივართ აპლიკაციიდან 
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

                // Open the copied Excel file and modify it using Interop
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(destinationFilePath);
                _Worksheet worksheet = workbook.Sheets[1]; // Assuming there is only one sheet
                Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

                // Find the column index for "Uniq_Id_New"
                int column = 1;
                while (worksheet.Cells[1, column].Value != null && worksheet.Cells[1, column].Value.ToString() != "Uniq_Id_New")
                {
                    column++;
                }

                if (worksheet.Cells[1, column].Value == null)
                {
                    // If "Uniq_Id_New" header does not exist, add it
                    worksheet.Cells[1, column].Value = "Uniq_Id_New";
                }

                // Find the column index for "Uniq_Id_MDB"
                while (worksheet.Cells[1, column+1].Value != null && worksheet.Cells[1, column+1].Value.ToString() != "Uniq_Id_MDB")
                {
                    column++;
                }

                if (worksheet.Cells[1, column + 1].Value == null)
                {
                    // If "Uniq_Id_MDB" header does not exist, add it
                    worksheet.Cells[1, column + 1].Value = "Uniq_Id_MDB";
                }

                // Read Excel data into a dictionary for faster lookup
                Dictionary<string, string> excelData = new Dictionary<string, string>();
                int rowCount = worksheet.UsedRange.Rows.Count;
                for (int i = 2; i <= rowCount; i++) // Assuming data starts from row 2
                {
                    string uniqId = worksheet.Cells[i, 1]?.Value?.ToString();
                    string litterId = worksheet.Cells[i, 2]?.Value?.ToString();

                    // Only add to dictionary if both uniqId and litterId are not null
                    if (!string.IsNullOrEmpty(uniqId) && !string.IsNullOrEmpty(litterId))
                    {
                        excelData.Add($"{uniqId}_{litterId}", worksheet.Cells[i, column]?.Value?.ToString() ?? "");
                    }
                }

                // Fetch data from database
                var dbData = geographicDynamicDbContext.QarsafariGroupeds.ToList();

                // Update Excel with fetched data
                for (int i = 2; i <= rowCount; i++)
                {
                    string excelUniqId = worksheet.Cells[i, 1]?.Value?.ToString();
                    string excelLitterId = worksheet.Cells[i, 2]?.Value?.ToString();

                    if (!string.IsNullOrEmpty(excelUniqId) && !string.IsNullOrEmpty(excelLitterId))
                    {
                        var matchedData = dbData.FirstOrDefault(item =>
                            item.UniqIdOld.ToString() == excelUniqId &&
                            item.LiterId.ToString() == excelLitterId);

                        if (matchedData != null)
                        {
                            // Write new uniqId from DB
                            worksheet.Cells[i, column].Value = matchedData.UniqId;

                            // Find corresponding MDB value safely
                            //var match = excelDataList
                            //    .FirstOrDefault(m => m.Uniq_ID_gadanomrili == matchedData.UniqId.ToString());

                            var match = excelDataList.FirstOrDefault(m =>
                                        m.Uniq_ID_gadanomrili == matchedData.UniqId.ToString() &&
                                        m.Litter_Id == matchedData.LiterId.ToString());

                            if (match != null)
                            {
                                worksheet.Cells[i, column + 1].Value = match.Uniq_Id_MDB;
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


                // Save changes and cleanup
                workbook.Save();
                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                result.Success = true;
                result.StatusCode = System.Net.HttpStatusCode.OK;
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occurred during file operations or Excel manipulation
                result.Success = false;
                result.StatusCode = System.Net.HttpStatusCode.InternalServerError;
                result.Message = $"ძირი ექსელის გადაკოპირებისა და მასში ახალი UniqId ჩაწერისას მოხდა შეცდომა. შეცდომის რიგი: {ex.Message}";
                // Optionally log the exception details for troubleshooting
            }

            return result;
        }


        // ამრგვალებს 5 ის ჯერადზე გადაცემულ რიცხვს
        static int RoundToNearest(double number)
        {
            int i = Convert.ToInt32(number);
            return (i % 5) == 0 ? i : (i % 5) >= 2.5 ? i + 5 - (i % 5) : i - (i % 5);
        }
        #endregion
    }
}
