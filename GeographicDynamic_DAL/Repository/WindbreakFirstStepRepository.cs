using Azure;
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
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static Azure.Core.HttpHeader;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using GeographicDynamic_DAL.Models.WindbreakFirstStepMethods;
//using static System.Net.Mime.MediaTypeNames;
using System.Drawing;
using System.Drawing.Imaging;
namespace GeographicDynamic_DAL.Repository
{
    public class WindbreakFirstStepRepository : IWindbreakFirstStep
    {
        WindbreakFirstStepMethods _windbreakMethods = new WindbreakFirstStepMethods();

        public Result<bool> GetCheckPhotoDate(string folderPath, string resultPath)
        {
            var testdb = "";
            List<string> UnMachedPhotos = new List<string>();
            try
            {
                string ForExcelName = Path.GetFileName(resultPath);
                var directories = Directory.GetDirectories(folderPath);

                //for proggress bar 
                int currentDirectoryIndex = 0;
                foreach (var Liters in directories)
                {

                    var actualPathOfFolderLitter = Liters.LastIndexOf('\\');
                    var numericalPathLitter = Liters.Substring(actualPathOfFolderLitter + 1);
                    var litterPath = Directory.GetDirectories(Liters);
                    int totalDirectories = directories.Length;
                    currentDirectoryIndex++;

                    foreach (var d in litterPath)
                    {
                        testdb = d;
                        var actualPathOfFolder = d.LastIndexOf('\\');
                        var numericalPath = d.Substring(actualPathOfFolder + 1);

                        //var directoriesWithoutExtention = Directory.GetDirectories(d).OrderBy(x => Convert.ToInt32(Path.GetFileNameWithoutExtension(x)));
                        var directoriesWithoutExtention = Directory.GetFiles(d);

                        //var list = directoriesWithoutExtention.OrderBy(y => int.Parse(y.Split('\\')[5])).ToList();



                        //foreach (var item in list)
                        //{

                        DirectoryInfo d5 = new DirectoryInfo(d);
                        FileInfo[] infos1 = d5.GetFiles();
                        var innerdir = d;
                        var files = Directory.GetFiles(d).Where(m => !m.Contains(".db") && (m.Contains(".jpg") || m.Contains(".jpeg"))).ToList();

                        int photoLength = 0;
                        bool photoebiaremtxveva = false;
                        string formattedDate;
                        using (Image myImageFirst = Image.FromFile(files[0]))
                        {

                            try
                            {

                                PropertyItem propItemFirst = myImageFirst.GetPropertyItem(306);
                                DateTime dtakenFirst;
                                //// ახლა ვიღებთ data taken ს პირველი ფოტო სურათის რათა შევადაროთ დანარჩენებს
                                string sdateFirst = Encoding.UTF8.GetString(propItemFirst.Value).Trim();
                                string secondhalfFirst = sdateFirst.Substring(sdateFirst.IndexOf(" "), (sdateFirst.Length - sdateFirst.IndexOf(" ")));
                                string firsthalfFirst = sdateFirst.Substring(0, 10);
                                firsthalfFirst = firsthalfFirst.Replace(":", "-");
                                sdateFirst = firsthalfFirst + secondhalfFirst;
                                DateTime firstFileDate = DateTime.Parse(sdateFirst);
                                formattedDate = firstFileDate.ToString("MM/dd/yy");
                            }
                            catch (ArgumentException)
                            {
                                // Date Taken property not found, fallback to modified date
                                FileInfo fileInfo = new FileInfo(files[0]);
                                formattedDate = fileInfo.LastWriteTime.ToString("MM/dd/yy");
                            }

                        }

                        foreach (var file in files)
                        {
                            using (Image myImage = Image.FromFile(file))
                            {
                                DateTime dtaken;

                                try
                                {

                                    PropertyItem propItem = myImage.GetPropertyItem(306);
                                    //Convert date taken metadata to a DateTime object
                                    string sdate = Encoding.UTF8.GetString(propItem.Value).Trim();
                                    string secondhalf = sdate.Substring(sdate.IndexOf(" "), (sdate.Length - sdate.IndexOf(" ")));
                                    string firsthalf = sdate.Substring(0, 10);
                                    firsthalf = firsthalf.Replace(":", "-");
                                    sdate = firsthalf + secondhalf;
                                    dtaken = DateTime.Parse(sdate);
                                }
                                catch (ArgumentException)
                                {
                                    // Date Taken property not found, fallback to modified date
                                    FileInfo fileInfo = new FileInfo(file);
                                    dtaken = fileInfo.LastWriteTime;
                                }


                                photoLength++;
                                var extensionTest = Path.GetExtension(file);

                                //DateTime fileDate = File.GetCreationTime(file);
                                //DateTime fileDate = File.GetLastWriteTime(file);



                                // string formattedDateToCompare = fileDate.ToString("MM/dd/yy");
                                string formattedDateToCompare = dtaken.ToString("MM/dd/yy");

                                if (formattedDateToCompare != formattedDate && !photoebiaremtxveva)
                                {
                                    UnMachedPhotos.Add(numericalPathLitter + " | " + "/" + numericalPath.ToString() + "/" + " |თარიღები არ ემთხვევა");
                                    photoebiaremtxveva = true;
                                    // MessageBox.Show("მოხდა შეცდომა ფაილების თარიღები არ ემთხვევა ერთმანეთს !");
                                }


                                //}

                            }
                        }
                        if (photoLength < 3)
                        {
                            UnMachedPhotos.Add(numericalPathLitter + " | " + "/" + numericalPath.ToString() + "/" + " |ნაკლები ფოტოა");
                            //MessageBox.Show("მოხდა შეცდომა ფოლდერში 3-ზე ნაკლები ფოტოს ფაილია აღმოჩენილი !");
                        }
                        // აქ ემატება შემოწმება იმისთვის რომ ინახოს თუ არის 10 ზე მეტი ფოტო ფილდერში და თუ კი მაშინ წაშალოს იმდენი რამდენითაც მეტია 10 ზე 
                        // წასაშლელად ირჩევა შემთხვევითი პრინციპით რიცხვი მასივში და შენდეგ იშლება
                        //if (photoLength > 10)
                        //{
                        //    Random random = new Random();
                        //    int filesToDeleteCount = photoLength - 10; // Calculate the number of files to delete

                        //    for (int i = 0; i < filesToDeleteCount; i++)
                        //    {
                        //        // Select a random file from the list
                        //        int randomIndex = random.Next(files.Count);
                        //        string fileToDelete = files[randomIndex];

                        //        // Delete the file
                        //        File.Delete(fileToDelete);

                        //        // Remove the file from the list
                        //        files.RemoveAt(randomIndex);
                        //    }

                        //    // Update photoLength after deletion
                        //    photoLength = files.Count;
                        //}


                    }
                }

                WriteToExcel(UnMachedPhotos, ForExcelName);

                void WriteToExcel(List<string> UnMatchedPhotos, string ForExcelName)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("ResultSheet");

                            //Writing data into excel
                            worksheet.Cell(1, "A").Value = "შეცდომები";
                            //worksheet.Cell(1, "B").Value = "UNIQ_ID";

                            for (int r = 0; r < UnMatchedPhotos.Count(); r++) //r stands for ExcelRow and c for ExcelColumn
                            {
                                string[] parts = UnMatchedPhotos[r].Split('/');
                                worksheet.Cell(r + 2, "A").Value = string.Concat(parts);
                                //worksheet.Cell(r + 2, "B").Value = parts[1];
                            }

                            workbook.SaveAs(resultPath + "\\Results-" + ForExcelName + ".xlsx");
                        }

                        //Process.Start(resultPath + "\\Results-" + ForExcelName + ".xlsx");
                        //Process.Start(new ProcessStartInfo { FileName = @"${resultPath}\\Results-{ForExcelName}.xlsx", UseShellExecute = true });
                    }

                    catch (Exception exHandle)

                    {

                        Console.WriteLine("Exception: " + exHandle.Message);

                        Console.ReadLine();

                    }
                }
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                var testdbdb = testdb;
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "მოხდა შეცდომა:" + ex.Message
                };
            }
        }
        //ფოტოების გაყოფის ფუნცქია 
        public Result<bool> PhotoSplitKerdzoSaxelmwifo(string GadanomriliPhotoFolderPath, string DestinationFolderPath)
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
            try
            {
                GadanomriliFotoebi photo = new GadanomriliFotoebi();
                var directories = Directory.GetDirectories(GadanomriliPhotoFolderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                foreach (var folderPath in directories)
                {
                    var idxLiter = folderPath.LastIndexOf('\\');
                    string literIDstr = folderPath.Substring(idxLiter + 1);

                    double literID = Convert.ToDouble(literIDstr);

                    var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                    var list = directories1.OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));


                    foreach (var item in list)
                    {
                        DirectoryInfo d5 = new DirectoryInfo(item);
                        FileInfo[] infos1 = d5.GetFiles();

                        var idxUniqid = item.LastIndexOf('\\');

                        string uniqIDstr = item.Substring(idxUniqid + 1);

                        double uniqID = Convert.ToDouble(uniqIDstr);

                        string photoN = "";

                        var PhotoDate = "";

                        photo.UniqId = uniqIDstr;
                        bool ismoved = true;
                        foreach (FileInfo f6 in infos1)
                        {
                            if (!f6.Name.Contains(".db"))
                            {

                                QarsafariGrouped? qarsafaritest = geographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(m => m.UniqId == uniqID);

                                // აქ გვჭირდება რომ მოწმდებოდეს მარტო კერძო ან სახელმწიფო რადგან ბაზაში იურიდიული პირიდა მუნიცპალიტეტი აღარაა მარტო კერძო ან სახელმწიფო
                                if (qarsafaritest?.Owner == "კერძო" || qarsafaritest?.Owner == "იურიდიული პირი")
                                {
                                    if (ismoved)
                                    {

                                        photo.LiterId = literID;
                                        string destinationFolder = Path.Combine((string.Concat(DestinationFolderPath + "\\" + "photoSplit" + "\\" + "Kerdzo")), literID.ToString());
                                        if (!Directory.Exists(destinationFolder))
                                        {
                                            Directory.CreateDirectory(destinationFolder);
                                        }
                                        string destinationFile = Path.Combine(destinationFolder, uniqID.ToString());
                                        //File.Copy(item, destinationFile);
                                        Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(item, destinationFile);

                                        ismoved = false;
                                    }

                                }
                                if (qarsafaritest?.Owner != "კერძო" && qarsafaritest?.Owner != "იურიდიული პირი")
                                {
                                    if (ismoved)
                                    {
                                        photo.LiterId = literID;


                                        string destinationFolder = Path.Combine((string.Concat(DestinationFolderPath + "\\" + "photoSplit" + "\\" + "Saxelmwifo")), literID.ToString());

                                        if (!Directory.Exists(destinationFolder))
                                        {
                                            Directory.CreateDirectory(destinationFolder);
                                        }
                                        string destinationFile = Path.Combine(destinationFolder, uniqID.ToString());
                                        //File.Copy(item, destinationFile);
                                        Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(item, destinationFile);
                                        ismoved = false;
                                    }
                                }
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "მოხდა შეცდომა:" + ex.Message
                };
            }





            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }

        //ფოტოების გადანომრვის ფუნქციონალი ეშვება აქ 
        public Result<bool> RenamePhotosInFolder(RenamePhotoDTO renamePhotoDTO)
        {
            #region GIORGI
            //var directories = Directory.GetDirectories(renamePhotoDTO.FolderPath).OrderBy(filePath => Convert.ToString(Path.GetFileNameWithoutExtension(filePath)));
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            geographicDynamicDbContext.GadanomriliFotoebis.ExecuteDelete();
            var directories = Directory.GetDirectories(renamePhotoDTO.FolderPath).OrderBy(filePath => int.Parse(Path.GetFileNameWithoutExtension(filePath)));

            //int foldercount = renamePhotoDTO.FolderStartNumber;
            int photocount = renamePhotoDTO.PhotoStartNumber;
            var random = new Random();
            var tempFolderCount = random.Next(100000, 999999);

            if (renamePhotoDTO.Gadanomrilia == false)
            {
                try
                {
                    #region აქ კეთდება წინასწარ რანდომ რიცხვის მიმატება ფოლდდერების სახელებს მანამ გადანომვრას დავიწყებთ
                    ////ფაილების გადანომვრა 
                    //foreach (var folderPath in directories)
                    //{                 //ფაილების გადანომვრა რენდომ რიცხვით რომ გამოირიცხოს დუპლიკატი
                    //    var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));
                    //    var list = directories1.OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));
                    //    foreach (var items in list)
                    //    {
                    //        var idx = items.LastIndexOf('\\');
                    //        string kk = items.Substring(idx + 1);

                    //        var newname = items.Replace(kk, Convert.ToString(foldercount)); //es mushaobs

                    //        int idx11 = items.LastIndexOf('\\');
                    //        string oldfoldername = items.Substring(0, idx11);
                    //        string newnamefolder = oldfoldername + "\\" + tempFolderCount; //ჯერ ეს უნდა გავუშვათ 

                    //        Directory.Move(items, newnamefolder);
                    //        tempFolderCount++;
                    //    }
                    //}
                    #endregion
                    //გადანომვრის ციკლი შერჩეული რიცხვით სადანაც გვინდა დაიწყოს 
                    List<Qarsafari> qarsafarisForRename = geographicDynamicDbContext.Qarsafaris.Where(m => m.IsUniqLiterNull == "true").ToList();
                    foreach (var item in qarsafarisForRename)
                    {
                        // ჯერ ვეძებთ არჩეულ ფოლდერში ლიტერიდი თი შესაბამის ფოლდერს
                        string folderLiterIDPath = Directory.GetDirectories(renamePhotoDTO.FolderPath, item.LiterId.ToString(), SearchOption.AllDirectories).FirstOrDefault();
                        // ჯერ ვეძებთ არჩეულ ფოლდერში უნიკიდი თი შესაბამის 
                        string folderUniqIDPath = Directory.GetDirectories(folderLiterIDPath, item.UniqIdOld.ToString(), SearchOption.AllDirectories).FirstOrDefault();
                        if (!string.IsNullOrEmpty(folderUniqIDPath)) //თუ იპოვა
                        {

                            int idx11 = folderUniqIDPath.LastIndexOf('\\');
                            string oldfoldername = folderUniqIDPath.Substring(0, idx11);
                            string newnamefolder = oldfoldername + "\\" + Convert.ToString(item.UniqId.ToString() + tempFolderCount);
                            Directory.Move(folderUniqIDPath, newnamefolder);

                        }
                        else // თუ ვერ იპოვა
                        {
                            return new Result<bool> { Success = false, StatusCode = System.Net.HttpStatusCode.OK, Message = "უნიკიდ" + item.UniqIdOld.ToString() + " ვერ მოიძებნა ფოლდერებში" };
                        }
                    }
                    // რომ დასრულდება ბოლოს ფოლდერის სახელებს უნდა ჩამოვაჭრათ ბოლო 6 სიმბოლო რაც წინასწარ დავუმატეთ
                    foreach (var folderPath in directories)
                    {                 //ფაილების გადანომვრა რენდომ რიცხვით რომ გამოირიცხოს დუპლიკატი
                        var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToDouble(Path.GetFileNameWithoutExtension(filePath)));
                        var list = directories1.OrderBy(filePath => Convert.ToDouble(Path.GetFileNameWithoutExtension(filePath)));
                        foreach (var items in list)
                        {

                            // Rename the current folder
                            string folderNameOld = Path.GetFileName(items);
                            string newFolderName = folderNameOld.Replace(tempFolderCount.ToString(), "");
                            if (newFolderName != folderNameOld)
                            {
                                string newFolderPath = Path.Combine(folderPath, newFolderName);

                                Directory.Move(items, newFolderPath);
                            }

                        }
                    }

                    //// ძველი გადანომვრა
                    //foreach (var folderPath in directories)
                    //{
                    //    int idx2 = folderPath.LastIndexOf('\\');
                    //    var kk2 = Convert.ToInt32(folderPath.Substring(idx2 + 1));

                    //    var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                    //    var list = directories1.OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                    //    foreach (var items in list)
                    //    {
                    //        var idx = items.LastIndexOf('\\');
                    //        string kk = items.Substring(idx + 1);

                    //        var newname = items.Replace(kk, Convert.ToString(foldercount)); //es mushaobs

                    //        int idx11 = items.LastIndexOf('\\');
                    //        string oldfoldername = items.Substring(0, idx11);
                    //        string newnamefolder = oldfoldername + "\\" + Convert.ToString(foldercount);
                    //        Directory.Move(items, newnamefolder);

                    //        foldercount++;
                    //    }
                    //}
                }
                catch (Exception ex) { }
            }
            //ფოტოების გადანომვრა 
            try
            {

                foreach (var folderPath in directories)
                {
                    var idxLiter = folderPath.LastIndexOf('\\');
                    string literIDstr = folderPath.Substring(idxLiter + 1);

                    double literID = Convert.ToDouble(literIDstr);

                    var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToDouble(Path.GetFileNameWithoutExtension(filePath)));
                    var list = directories1.OrderBy(filePath => Convert.ToDouble(Path.GetFileNameWithoutExtension(filePath)));

                    foreach (var item in list)
                    {
                        DirectoryInfo d5 = new DirectoryInfo(item);

                        FileInfo[] infos1 = d5.GetFiles();

                        var idxUniqid = item.LastIndexOf('\\');
                        string uniqIDstr = item.Substring(idxUniqid + 1);

                        string photoN = "";

                        var PhotoDate = "";
                        string photoNCorrected = "";
                        // ფოტოების გადასანომრი ციკლი 
                        foreach (FileInfo f6 in infos1)
                        {
                            if (!f6.Name.Contains(".db"))
                            {
                                var ext = Path.GetExtension(f6.FullName);
                                var newPhotoNamePath = f6.FullName.Replace(f6.Name, Convert.ToString(photocount) + ext);
                                File.Move(f6.FullName, newPhotoNamePath);

                                photoN += Convert.ToString(photocount) + "/";

                                //ფოტოს თარიღის წამოღება
                                bool isWritten = false;
                                if (!isWritten)
                                {
                                    DateTime photoDate;
                                    DateTime dtaken;
                                    try
                                    {
                                        // Try to get the Date Taken property
                                        using (Image myImage = Image.FromFile(@newPhotoNamePath))
                                        {
                                            PropertyItem propItem = myImage.GetPropertyItem(306);
                                            string sdate = Encoding.UTF8.GetString(propItem.Value).Trim();
                                            string secondhalf = sdate.Substring(sdate.IndexOf(" "), (sdate.Length - sdate.IndexOf(" ")));
                                            string firsthalf = sdate.Substring(0, 10);
                                            firsthalf = firsthalf.Replace(":", "-");
                                            sdate = firsthalf + secondhalf;
                                            dtaken = DateTime.Parse(sdate, CultureInfo.InvariantCulture);
                                        }
                                    }
                                    catch (ArgumentException)
                                    {
                                        // Date Taken property not found, fallback to modified date
                                        dtaken = f6.LastWriteTime;
                                    }
                                    ////var modifiedDate1 = f6.LastWriteTime;
                                    ////ვიღებთ ფოტოს data taken-ს modify თარიღის ნაცვლად
                                    //Image myImage = Image.FromFile(@newPhotoNamePath);
                                    //PropertyItem propItem = myImage.GetPropertyItem(306);
                                    //DateTime dtaken;

                                    ////Convert date taken metadata to a DateTime object
                                    //string sdate = Encoding.UTF8.GetString(propItem.Value).Trim();
                                    //string secondhalf = sdate.Substring(sdate.IndexOf(" "), (sdate.Length - sdate.IndexOf(" ")));
                                    //string firsthalf = sdate.Substring(0, 10);
                                    //firsthalf = firsthalf.Replace(":", "-");
                                    //sdate = firsthalf + secondhalf;
                                    //dtaken = DateTime.Parse(sdate);

                                    var formatInfo = new CultureInfo("en-US").DateTimeFormat;
                                    formatInfo.DateSeparator = "-";
                                    //PhotoDate = modifiedDate1.ToString("dd-MM-yyyy", formatInfo);
                                    //ვიღებთ ფოტოს data taken-ს modify თარიღის ნაცვლად
                                    PhotoDate = dtaken.ToString("dd-MM-yyyy", formatInfo);
                                }
                                isWritten = true;
                                photocount++;

                            }

                        }
                        // SQL ბაზაში დამატება და ცვლილებების დამახსოვრება 
                        GadanomriliFotoebi photo = new GadanomriliFotoebi();
                        photo.UniqId = uniqIDstr;
                        photo.LiterId = literID;
                        photoNCorrected = photoN.TrimEnd('/'); // ბოლოში სლექშებს უშლის 
                        photo.PhotoN = photoNCorrected;
                        photo.PhotoDate = PhotoDate;
                        geographicDynamicDbContext.GadanomriliFotoebis.Add(photo); // ამით ემატება ბაზაში 
                        geographicDynamicDbContext.SaveChanges();//ამით ამახსოვრებს 
                    }
                }
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK, Message = "წარმატებით გადაინომრა" };
            }
            catch (Exception ex)
            {
                return new Result<bool> { Success = false, StatusCode = System.Net.HttpStatusCode.OK, Message = ex + "ჩტო ტა ნიტო" };
            }
            #endregion


            return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
        }



        //ექსელის ფაილებთან სამუშაო ფუნქცია სადაც უნდა წამოვიდეს უშუალოდ ექსელის წაკითხვისას 
        public Result<bool> ExcelCalculations(ExcelReadDTO excelReadDTO)

        {

            ////////////var createResultFolder = _windbreakMethods.copyOldExcelOriginal(excelReadDTO);



            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            ////////// ჩეკბოქსების გამოტანა ცვლადში რომ ადვილად აიწყოს შემოწმება 
            var CalcVarjisFartiCheckbox = excelReadDTO.CalcVarjisFartiCheckbox;
            var AccessShitNameTextbox = excelReadDTO.AccessShitName;


            /////////ფუნქციების გამოძახებები თავის შეცდომიანად თუ სადმე რამე იყო 
            //// ვკითხულობთ ექსელიდან ინფორმაციას და შეგვაქვს sql ში qarsafari ცხრილი
            var ExcelisWakitxvaRestult = _windbreakMethods.ExcelisWakitxva(excelReadDTO);
            if (ExcelisWakitxvaRestult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ექსელის წაკითხვის დროს" + ExcelisWakitxvaRestult.Message
                };
            }

            var fillAmountOfSpecesResult = _windbreakMethods.fillAmountOfSpeces(excelReadDTO);
            if (fillAmountOfSpecesResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "არიქნაა" + fillAmountOfSpecesResult.Message
                };
            }

            List<Qarsafari> newQarsafari = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).ToList();
            //ვამოწმებთ Excel-ში თუ არის დუპლიკატი Unic-Liter - ID - ები
            var ShemowmebaUnicLiterExcelshiResult = _windbreakMethods.ShemowmebaUnicLiterExcelshi();
            if (ShemowmebaUnicLiterExcelshiResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა Access ფაილის წაკითხვისას" + ShemowmebaUnicLiterExcelshiResult.Message
                };
            }



            //// ვკითხულობთ აქსესიდან ინფორმაციას და შეგვაქვს sql ში WindbreakMDB ცხრილი
            if (AccessShitNameTextbox != "")
            {
                var AccessReadingResult = _windbreakMethods.AccessWakitxva(excelReadDTO.AccessFilePath, excelReadDTO.AccessShitName);

                if (AccessReadingResult.Success == false)
                {
                    return new Result<bool>
                    {
                        Success = false,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = "მოხდა შეცდომა Access ფაილის წაკითხვისას" + AccessReadingResult.Message
                    };
                }
            }

            ///////////////////////ეს ფუქნცია ამოწმებს excel და access ცხრილებს და ადარებს UNIQID ებს თუ ემთხვევა ერთმანეთს
            var ShemowmebaAccessExcelUnicLiterDublicatsResult = _windbreakMethods.ShemowmebaAccessExcelUnicLiterDublicats();
            if (ShemowmebaAccessExcelUnicLiterDublicatsResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = ShemowmebaAccessExcelUnicLiterDublicatsResult.Message + " ზედმეტია: " + ShemowmebaAccessExcelUnicLiterDublicatsResult.Data[0]
                };
            }

            if (CalcVarjisFartiCheckbox == true)
            {
                var ChaweraVarjisPartiFunction = _windbreakMethods.ChaweraVarjisParti(excelReadDTO.ProjectNameID);

                if (ChaweraVarjisPartiFunction.Success == false)
                {
                    return new Result<bool>
                    {
                        Success = false,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = "ვარჯის ფართის დათვლის დროს მოხდა შეცდომა" + ChaweraVarjisPartiFunction.Message
                    };
                }
            }
            var CheckerOfVarjisFartiandSaxeobaResult = _windbreakMethods.CheckerOfVarjisFartiandSaxeoba();
            if (CheckerOfVarjisFartiandSaxeobaResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ვარჯის ფართის ჩაწერისას " + CheckerOfVarjisFartiandSaxeobaResult.Message
                };
            }
            var UIDReplaceExcelResult = _windbreakMethods.UIDReplaceExcel();
            if (UIDReplaceExcelResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ექსელში UID-ის ჩაწერის დროს" + UIDReplaceExcelResult.Message
                };
            }


            var QarsafariGadanomrvaResult = _windbreakMethods.QarsafariGadanomrva(excelReadDTO.UnicIDStartNumber);
            if (QarsafariGadanomrvaResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = QarsafariGadanomrvaResult.Message
                };
            }

            ///////////// ჯერ არ ვიყიენებთ მარა გამოსაყენებელია ხეხილში ვამოწმებთ დუბლიკატები ხომ არ არის
            var QarsafariXexilisShemowmebaResult = _windbreakMethods.QarsafariXexilisShemowmeba();
            if (QarsafariXexilisShemowmebaResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = QarsafariXexilisShemowmebaResult.Message
                };
            }

            //aq unda fotoebi renamephotos
            if (excelReadDTO.GadanomriliaFotoebi != true)
            {
                RenamePhotoDTO renamePhotoDTOForRename = new RenamePhotoDTO();
                renamePhotoDTOForRename.FolderPath = excelReadDTO.FolderPath;
                renamePhotoDTOForRename.PhotoStartNumber = excelReadDTO.PhotoStartNumber;
                var RenamePhotosInFolderFromExcel = RenamePhotosInFolder(renamePhotoDTOForRename);
                if (RenamePhotosInFolderFromExcel.Success == false)
                {
                    return new Result<bool>
                    {
                        Success = false,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = RenamePhotosInFolderFromExcel.Message
                    };
                }
            }
            //aq unda fotoebi renamephotos

            var ProcentisDatvlaResult = _windbreakMethods.QarsafariProcentisDatvla();
            if (ProcentisDatvlaResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = ProcentisDatvlaResult.Message
                };
            }

            // გადაგვაქვს ინფორმაცია აქსესიდან ექსელში
            var UpdateFromAccessToExcellResult = _windbreakMethods.UpdateFromAccessToExcell();
            if (UpdateFromAccessToExcellResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = UpdateFromAccessToExcellResult.Message
                };
            }

            // შევსება ველების სადაც ვინახავთ რომელი მუნიციპალიტეტია და რომელი ეტაპია დათვლის 
            var FillProjectEtapiIDSResult = _windbreakMethods.FillProjectEtapiIDS(excelReadDTO.ProjectNameID, excelReadDTO.EtapiID);
            if (FillProjectEtapiIDSResult.Success = false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა პროექტის სახელის და ეტაპის მინიჭების დროს "
                };
            }
            //საკუთრებაში ვწერთ სახელმწიფოა თუ კერძო
            var FillSakutrebaIsKerdzoOrSaxelmwifoResult = FillSakutrebaIsKerdzoOrSaxelmwifo();
            if (FillSakutrebaIsKerdzoOrSaxelmwifoResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა საკუთრების მინიჭების დროს "
                };
            }

            //var UIDReplaceAccessResult = _windbreakMethods.UIDReplaceAccess();
            //if (UIDReplaceAccessResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = UIDReplaceAccessResult.Message
            //    };
            //}


            ////////////axali funqcia UIDREPLACE () {} // table qarsafarshi
            ////////////SET UID = str([ლიტერი ID]) + str([უნიკ ID]) // str chventan aris Convert.ToString()
            ///////////////////// ჯერ არ ვიყიენებთ მარა გამოსაყენებელია ხეხილში ვამოწმებთ დუბლიკატები ხომ არ არის
            ////////var QarsafariXexilisShemowmebaResult = _windbreakMethods.QarsafariXexilisShemowmeba();
            ////////if (QarsafariXexilisShemowmebaResult.Success == false)
            ////////{
            ////////    return new Result<bool>
            ////////    {
            ////////        Success = false,
            ////////        StatusCode = System.Net.HttpStatusCode.BadGateway,
            ////////        Message = QarsafariXexilisShemowmebaResult.Message
            ////////    };
            ////////}



            // qarsafari ცხრილის დაგრუპვა uniqid ის მიხედვით და გადატანა qarsafariGrouped ში
            var QarsafariToQarsafariGroupedResult = _windbreakMethods.QarsafariToQarsafariGrouped();
            if (QarsafariToQarsafariGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = QarsafariToQarsafariGroupedResult.Message
                };
            }



            var UIDReplaceQarsafariGroupedResult = _windbreakMethods.UIDReplaceQarsafariGrouped();
            if (UIDReplaceQarsafariGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = UIDReplaceQarsafariGroupedResult.Message
                };
            }


            var GadanomriliFotoebiToQarsafariGroupedResult = _windbreakMethods.GadanomriliFotoebiToQarsafariGrouped();
            if (GadanomriliFotoebiToQarsafariGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = GadanomriliFotoebiToQarsafariGroupedResult.Message
                };
            }

            #region ეს აღარ გვჭირდება რადგან არ მუშაობს 
            //var UPDTFromExcelToAccessResult = _windbreakMethods.UPDTFromExcelToAccess(excelReadDTO.AccessShitName);
            //if (UPDTFromExcelToAccessResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = UPDTFromExcelToAccessResult.Message
            //    };
            //}
            #endregion

            var UpdateFromQarsafariGroupedToAccessFileResult = _windbreakMethods.UpdateFromQarsafariGroupedToAccessFile(excelReadDTO.AccessShitName, excelReadDTO.AccessFilePath);
            if (UpdateFromQarsafariGroupedToAccessFileResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = UpdateFromQarsafariGroupedToAccessFileResult.Message
                };
            }
            /////////// ქარსაფარი გრუპდის ცხრილები რომ ამოექსპორტდეს 
            List<QarsafariGrouped> qarsafariGroupeds = geographicDynamicDbContext.QarsafariGroupeds.OrderBy(m => m.UniqId).ToList();
            /////////List<QarsafariGrouped> qarsafariGroupedsSaxelmwifo = geographicDynamicDbContext.QarsafariGroupeds.Where(x => x.Sakutreba == "სახელმწიფო" || x.Sakutreba == "მუნიციპალიტეტი").OrderBy(m => m.UniqId).ToList();
            /////////List<QarsafariGrouped> qarsafariGroupedsKerdzo = geographicDynamicDbContext.QarsafariGroupeds.Where(x => x.Sakutreba == "კერძო" || x.Sakutreba == "იურიდიული პირი").OrderBy(m => m.UniqId).ToList();



            var WriteToExcelGroupedResult = _windbreakMethods.WriteToExcelGrouped(qarsafariGroupeds, excelReadDTO.ExcelDestinationPath, "QarsafariGrouped-" + DateTime.Now.ToString("yyyy-MM-dd"), excelReadDTO.AccessFilePath, excelReadDTO.AccessShitName);
            if (WriteToExcelGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = WriteToExcelGroupedResult.Message,
                };
            }
            ///////////////ფუნქციის გამოძახებები
            var copyOriginalExcelResult = _windbreakMethods.copyOldExcelOriginal(excelReadDTO);
            if (copyOriginalExcelResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = copyOriginalExcelResult.Message,
                };
            }

            ///////aqedan da zeviT 

            ////////////////////ეს არააა საჭირო 
            //var WriteToExcelOldRootWithNewUniqId = _windbreakMethods.WriteToExcelRootOne(newQarsafari, excelReadDTO.ExcelDestinationPath);
            //if (WriteToExcelOldRootWithNewUniqId.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = WriteToExcelOldRootWithNewUniqId.Message
            //    };
            //}


            #region didiqarsafaris amoeqsporteba
            //var WriteToExcelGroupedSaxelmwifoResult = _windbreakMethods.WriteToExcelGrouped(qarsafariGroupedsSaxelmwifo, excelReadDTO.ExcelDestinationPath, "QarsafariGrouped-Saxelmwifo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            //if (WriteToExcelGroupedSaxelmwifoResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = WriteToExcelGroupedSaxelmwifoResult.Message,
            //    };
            //}
            //var WriteToExcelGroupedKerdzoResult = _windbreakMethods.WriteToExcelGrouped(qarsafariGroupedsKerdzo, excelReadDTO.ExcelDestinationPath, "QarsafariGrouped-Kerdzo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            //if (WriteToExcelGroupedKerdzoResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = WriteToExcelGroupedKerdzoResult.Message,
            //    };
            //}

            /////////////////////ქარსაფარის ცხრილები რომ ამოექსპორტდეს 

            //List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).ToList();//მთლიანი ცხრილი 
            //List<Qarsafari> qarsafarisSaxelmwifo = geographicDynamicDbContext.Qarsafaris.Where(x => x.Sakutreba == "სახელმწიფო" || x.Sakutreba == "მუნიციპალიტეტი").ToList();//სახელმწიფო საკუთრების ცხრილი 
            //List<Qarsafari> qarsafarisKerdzo = geographicDynamicDbContext.Qarsafaris.Where(x => x.Sakutreba == "კერძო" || x.Sakutreba == "იურიდიული პირი").ToList();// კერძო საკუთრების ცხრილი 
            ////////////////////// ფუნქციის გამოძახებები 

            //var WriteToExcelQarsafariResult = _windbreakMethods.WriteToExcel(qarsafaris, excelReadDTO.ExcelDestinationPath, "Qarsafari-" + DateTime.Now.ToString("yyyy-MM-dd"));
            //if (WriteToExcelQarsafariResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = WriteToExcelQarsafariResult.Message,
            //    };
            //}
            //var WriteToExcelQarsafariSaxelmwifoResult = _windbreakMethods.WriteToExcel(qarsafarisSaxelmwifo, excelReadDTO.ExcelDestinationPath, "Qarsafari-Saxelmwifo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            //if (WriteToExcelQarsafariSaxelmwifoResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = WriteToExcelQarsafariSaxelmwifoResult.Message,
            //    };
            //}
            //var WriteToExcelQarsafariKerdzoResult = _windbreakMethods.WriteToExcel(qarsafarisKerdzo, excelReadDTO.ExcelDestinationPath, "Qarsafari-Kerdzo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            //if (WriteToExcelQarsafariKerdzoResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = WriteToExcelQarsafariKerdzoResult.Message,
            //    };
            //}
            #endregion

            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };

        }



        // აქ ხდება ფუნქციების აღწერა 



        // გიოს ნახლაფორთარი რომელიც ერთ ხაზში დაიწერა საკუთრების ველის შევსებისთვის 
        #region ნახლაფორთარი გიოსი 
        public Result<bool> FillSakutrebaIsKerdzoOrSaxelmwifo() // ფუნქცია გამოიყენება იმისთვის რომ საკუთრების ველში შეივსოს სახელწმიფოა თუ კერძო 
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            try
            {
                List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.ToList();

                string Sakutrebastore = "";// ინახება მნიშვნელობა როცა isUniqIdLiterIdtrue 
                foreach (var item in qarsafaris)
                {
                    if (item.IsUniqLiterNull == "true") // && item.Owner != null
                    {
                        if (item.Owner == "მუნიციპალიტეტი" || item.Owner == "სახელმწიფო" || String.IsNullOrEmpty(item.Owner))
                        {
                            item.Sakutreba = "სახელმწიფო";
                            Sakutrebastore = "სახელმწიფო";
                        }
                        else
                        {
                            item.Sakutreba = "კერძო";
                            Sakutrebastore = "კერძო";
                        }
                    }

                    if (item.IsUniqLiterNull == "false") // && item.Owner == null
                    {
                        item.Sakutreba = Sakutrebastore;
                    }


                    geographicDynamicDbContext.SaveChanges();
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
        #endregion






        // ეშვება GET ფუნქცია რომ მიიღოს მუნიციპალიტეტების სია FRONT-ში 
        public Result<DictionaryDTO> GetProjectNames()
        {
            try
            {

                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                var list = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 2).ToList();

                List<DictionaryDTO> DictionaryDTOs = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 2).Select(x => new DictionaryDTO { ID = x.Id, Name = x.Name }).ToList();

                return new Result<DictionaryDTO>
                {
                    Success = true,
                    Data = DictionaryDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დაბრუნდა სია"
                };
            }
            catch (Exception ex)
            {
                return new Result<DictionaryDTO>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "შეცდომა მოხდა " + ex.Message
                };

            }


        }
        //ეშვება GET გუნქცია რომ მიიღოს ეტაპის სია FRont-ში
        public Result<DictionaryDTO> GetEtapiID()
        {
            try
            {

                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                var list = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 3).ToList();

                List<DictionaryDTO> DictionaryDTOs = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 3).Select(x => new DictionaryDTO { ID = x.Id, Name = x.Name }).ToList();


                return new Result<DictionaryDTO>
                {
                    Success = true,
                    Data = DictionaryDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დაბრუნდა  ეტაპების სია"
                };
            }
            catch (Exception ex)
            {
                return new Result<DictionaryDTO>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "შეცდომა მოხდა " + ex.Message
                };

            }
        }

        // ეშვება GET  ფუნქცია რომ წამოიღოს ვარჯის ფართები მუნიციპალიტეტის მიხედვით ,
        public Result<VarjisFartiDTO> GetVarjisFartebi(int AreaNameID)
        {
            try
            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
                List<VarjisFartiDTO> VarjisFartiDTOs = geographicDynamicDbContext.VarjisFartis.Where(x => x.AreaNameId == AreaNameID).Select(x => new VarjisFartiDTO { Id = x.Id, Name = x.Saxeoba.Name, SaxeobaId = x.SaxeobaId, AreaNameId = x.AreaNameId, VarjisFarti1 = x.VarjisFarti1 }).ToList();
                return new Result<VarjisFartiDTO>
                {
                    Success = true,
                    Data = VarjisFartiDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დაბრუნდა სია"
                };
            }
            catch (Exception ex)
            {
                return new Result<VarjisFartiDTO>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "შეცდომა მოხდა " + ex.Message
                };
            }
        }

        public Result<DictionaryDTO> getSaxeobaList()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();



            List<DictionaryDTO> saxeobaList = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 1).Select(x => new DictionaryDTO { ID = x.Id, Name = x.Name, Code = x.Code }).OrderBy(x => x.Name).ToList();
            //saxeobaList.OrderByDescending(x=>x.Name);

            return new Result<DictionaryDTO>
            {
                Success = true,
                Data = saxeobaList,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }



        //public Result<bool> GroupingFormulas()
        //{
        //    GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
        //    List<Qarsafari> qarsafari = geographicDynamicDbContext.Qarsafaris.ToList();
        //    List<QarsafariGrouped> qarsafariGrouped = geographicDynamicDbContext.QarsafariGroupeds.ToList();
        //    try
        //    {
        //        foreach( var item in qarsafari)
        //        {

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
        //            Message = "მოხდა შეცდომა დაგრუპვის ფორმულის დროს" + ex.Message
        //        };
        //    }
        //}


        // ფუნცქია კითხულობს ბაზას და ქმნის ახალ ექსელის ფაილს რომ შევიდეს მონაცემები შიგნით და გაკეთდეს ექსელის დიდი ფაილი მხოლოდ ქარსაფარისთვის 



    }



}
