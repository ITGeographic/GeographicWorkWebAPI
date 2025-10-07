using ClosedXML.Excel;
using GeographicDynamic_DAL.Interface;
using GeographicDynamicWebAPI.Wrappers;
using System.Text;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;

namespace GeographicDynamic_DAL.Repository
{
    public class AeroImagesFolderingRepository : IAeroImagesFolderingRepository
    {

        #region დაფოლდერების პირველი მცდელობა წარუმატებელი 



        //public Result<bool> DafolderebaPotoebis()
        //{
        //    string excelPath = @"\\server\Urban\2025\Hotspots_Georgia_2025\SENT\2025_09_26_Alex_cxrili\Kazbegi_Akhmeta_20250923_Copy.xlsx";
        //    string imagesSourcePath = @"\\giswebserver\HotSpo - -Pictures\2025_09_23\AllPhotoesCopied";
        //    string videosSourcePath = @"\\giswebserver\HotSpo - -Pictures\2025_09_23\ALLVideos";
        //    string resultBasePath = @"\\giswebserver\HotSpo - -Pictures\2025_09_23\Dafolderebuli";

        //    try
        //    {
        //        using (var workbook = new XLWorkbook(excelPath))
        //        {
        //            var worksheet = workbook.Worksheet(1);
        //            var rows = worksheet.RangeUsed().RowsUsed();

        //            foreach (var row in rows.Skip(1))
        //            {
        //                string photoFileName = row.Cell(1).GetString().Trim().Replace("//", @"\");
        //                string folderName = row.Cell(2).GetString().Trim();

        //                if (string.IsNullOrWhiteSpace(photoFileName) || string.IsNullOrWhiteSpace(folderName))
        //                    continue;

        //                string sourcePhotoPath = Path.Combine(imagesSourcePath, photoFileName);
        //                string destFolder = Path.Combine(resultBasePath, folderName);
        //                Directory.CreateDirectory(destFolder);
        //                string destPath = Path.Combine(destFolder, photoFileName);

        //                if (!File.Exists(sourcePhotoPath))
        //                {
        //                    Console.WriteLine($"File not found: {sourcePhotoPath}");
        //                    continue;
        //                }

        //                // Retry copy with lock check
        //                if (!TrySafeCopy(sourcePhotoPath, destPath))
        //                {
        //                    Console.WriteLine($"Skipping locked file: {sourcePhotoPath}");
        //                    continue;
        //                }
        //            }
        //        }

        //        // Process videos
        //        var videoResult = SortVideosByImageDate(resultBasePath, videosSourcePath);
        //        if (!videoResult.Success)
        //        {
        //            Console.WriteLine("Video sorting failed!");
        //            return new Result<bool> { Success = false };
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error in DafolderebaPotoebis: {ex.Message}");
        //        return new Result<bool> { Success = false };
        //    }

        //    Console.WriteLine("Process finished!");
        //    return new Result<bool> { Success = true };
        //}

        //public Result<bool> SortVideosByImageDate(string resultBasePath, string videosSourcePath)
        //{
        //    try
        //    {
        //        var imageDirectories = Directory.GetDirectories(resultBasePath);

        //        var allVideos = Directory.GetFiles(videosSourcePath, "*.*", SearchOption.TopDirectoryOnly)
        //                                 .Where(f => f.EndsWith(".mp4", System.StringComparison.OrdinalIgnoreCase)
        //                                          || f.EndsWith(".avi", System.StringComparison.OrdinalIgnoreCase)
        //                                          || f.EndsWith(".mov", System.StringComparison.OrdinalIgnoreCase)
        //                                          || f.EndsWith(".mkv", System.StringComparison.OrdinalIgnoreCase))
        //                                 .ToList();

        //        foreach (var dir in imageDirectories)
        //        {
        //            var images = Directory.GetFiles(dir, "*.*", SearchOption.TopDirectoryOnly)
        //                                  .Where(f => f.EndsWith(".jpg", System.StringComparison.OrdinalIgnoreCase)
        //                                           || f.EndsWith(".jpeg", System.StringComparison.OrdinalIgnoreCase)
        //                                           || f.EndsWith(".png", System.StringComparison.OrdinalIgnoreCase))
        //                                  .ToList();

        //            if (!images.Any())
        //                continue;

        //            // Read EXIF safely
        //            var imageTimes = images.Select(img => GetDateTakenFromImage(img)).ToList();
        //            DateTime minTime = imageTimes.Min();
        //            DateTime maxTime = imageTimes.Max();

        //            foreach (var video in allVideos)
        //            {
        //                DateTime videoTime = GetDateFromVideo(video);

        //                if (videoTime >= minTime && videoTime <= maxTime)
        //                {
        //                    string destPath = Path.Combine(dir, Path.GetFileName(video));
        //                    if (!File.Exists(destPath))
        //                    {
        //                        TrySafeCopy(video, destPath);
        //                    }
        //                }
        //            }
        //        }

        //        return new Result<bool> { Success = true };
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error in SortVideosByImageDate: {ex.Message}");
        //        return new Result<bool> { Success = false };
        //    }
        //}

        //private DateTime GetDateTakenFromImage(string path)
        //{
        //    try
        //    {
        //        // Use Image.FromFile which opens/closes file safely
        //        using (var img = Image.FromFile(path))
        //        {
        //            const int PropertyTagExifDTOrig = 0x9003;
        //            if (img.PropertyIdList.Contains(PropertyTagExifDTOrig))
        //            {
        //                var propItem = img.GetPropertyItem(PropertyTagExifDTOrig);
        //                string dateTakenStr = Encoding.ASCII.GetString(propItem.Value).Trim('\0');
        //                if (DateTime.TryParseExact(dateTakenStr, "yyyy:MM:dd HH:mm:ss", null,
        //                    System.Globalization.DateTimeStyles.None, out DateTime dt))
        //                    return dt;
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        // fallback silently
        //    }

        //    return File.GetCreationTime(path); // fallback
        //}

        //private DateTime GetDateFromVideo(string path)
        //{
        //    return File.GetCreationTime(path);
        //}

        //private bool TrySafeCopy(string source, string dest, int maxRetries = 5, int delayMs = 200)
        //{
        //    for (int attempt = 1; attempt <= maxRetries; attempt++)
        //    {
        //        try
        //        {
        //            File.Copy(source, dest, true);
        //            Console.WriteLine($"Copied: {source} -> {dest}");
        //            return true;
        //        }
        //        catch (IOException ex) when ((ex.HResult & 0xFFFF) == 32) // sharing violation
        //        {
        //            Console.WriteLine($"Attempt {attempt} failed (locked): {source}");
        //            Thread.Sleep(delayMs);
        //        }
        //        catch (IOException ex)
        //        {
        //            Console.WriteLine($"Attempt {attempt} failed: {ex.Message}");
        //            Thread.Sleep(delayMs);
        //        }
        //    }

        //    return false; // skip if still locked
        //}




        #endregion

        // Update these paths as needed
        private readonly string imagesSourcePath = @"\\giswebserver\HotSpo - -Pictures\2025_09_23\AllPhotoesCopied";
        private readonly string resultBasePath = @"\\giswebserver\HotSpo - -Pictures\2025_09_23\Dafolderebuli";
        private readonly string excelPath = @"\\server\Urban\2025\Hotspots_Georgia_2025\SENT\2025_09_26_Alex_cxrili\Kazbegi_Akhmeta_20250923_Copy.xlsx";
        public Result<bool> DafolderebaPotoebis()
        {
            try
            {
                if (!File.Exists(excelPath))
                    return new Result<bool> { Success = false, Message = $"Excel file not found: {excelPath}" };

                var lines = File.ReadAllLines(excelPath);
                var skippedFiles = new List<string>();

                // Skip header row
                for (int i = 1; i < lines.Length; i++)
                {
                    var parts = lines[i].Split(',');
                    if (parts.Length < 2) continue;

                    string imageFileName = parts[0].Trim();
                    string folderName = parts[1].Trim();

                    if (string.IsNullOrEmpty(imageFileName) || string.IsNullOrEmpty(folderName))
                        continue;

                    string sourceFile = Path.Combine(imagesSourcePath, imageFileName);
                    string destFolder = Path.Combine(resultBasePath, folderName);
                    string destFile = Path.Combine(destFolder, imageFileName);

                    try
                    {
                        if (!File.Exists(sourceFile))
                        {
                            skippedFiles.Add(imageFileName);
                            Console.WriteLine($"Source file not found: {sourceFile}");
                            continue;
                        }

                        Directory.CreateDirectory(destFolder);

                        // Copy with overwrite
                        File.Copy(sourceFile, destFile, overwrite: true);
                    }
                    catch (IOException ex)
                    {
                        skippedFiles.Add(imageFileName);
                        Console.WriteLine($"Failed to copy {imageFileName}: {ex.Message}");
                    }
                }

                if (skippedFiles.Count > 0)
                {
                    Console.WriteLine("Skipped files:");
                    foreach (var file in skippedFiles)
                        Console.WriteLine(file);
                }

                return new Result<bool>
                {
                    Success = true,
                    Message = "Images copied successfully."
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    Message = $"Error: {ex.Message}"
                };
            }
        }



    }
}
