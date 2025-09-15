using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using GeographicDynamic_DAL.Interface;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Diagnostics;

namespace GeographicDynamic_DAL.Repository
{
    public class AeroGadagebaRepository : IAero
    {
        public Result<List<AeroRecord>> ExcelisWakiTxvaAero()
        {
            List<AeroRecord> ExcelInfo = new List<AeroRecord>();

            try
            {
                string ExcelPath = @"D:\Projects\2025\QarsaffariDatvlebi\TestAero\GPS_pnt_time.xlsx";

                using (var workbook = new XLWorkbook(ExcelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed().Skip(1); // skip header row

                    foreach (var row in rows)
                    {
                        string latStr = row.Cell(1).GetValue<string>();
                        string lonStr = row.Cell(2).GetValue<string>();
                        string timeStr = row.Cell(3).GetValue<string>();

                        DateTime? dataTaken = null;
                        if (!string.IsNullOrWhiteSpace(timeStr))
                        {
                            if (DateTime.TryParse(timeStr, out var parsed))
                                dataTaken = parsed;
                        }

                        ExcelInfo.Add(new AeroRecord
                        {
                            Latitude = latStr,
                            Longitude = lonStr,
                            DataTaken = dataTaken
                        });
                    }
                }

                // ===== New Logic =====
                OrganizeImagesByEarliest(ExcelInfo, @"D:\Projects\2025\QarsaffariDatvlebi\TestAero\images");

                return new Result<List<AeroRecord>>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                };
            }
            catch (Exception ex)
            {
                return new Result<List<AeroRecord>>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "ექსელის წაკითხვა ვერ მოხერხდა: " + ex.Message
                };
            }
        }

        private void OrganizeImagesByEarliest(List<AeroRecord> excelData, string imagePath)
        {
            while (true)
            {
                // Refresh remaining files
                var allFiles = Directory.GetFiles(imagePath, "*.*", SearchOption.TopDirectoryOnly).ToList();
                if (allFiles.Count == 0) break;

                // Step 1: find earliest image
                var earliestFile = allFiles.OrderBy(f => File.GetLastWriteTime(f)).First();
                DateTime earliestDate = File.GetLastWriteTime(earliestFile);

                // Step 2: adjust time to Georgian local (+4h offset correction)
                DateTime adjustedDate = earliestDate.AddHours(-4);

                // Step 3: find Excel match
                //var match = excelData.FirstOrDefault(r =>
                //    r.DataTaken.HasValue &&
                //    Math.Abs((r.DataTaken.Value - adjustedDate).TotalSeconds) <= 3);
                var match = excelData
                    .Where(r => r.DataTaken.HasValue)
                    .OrderBy(r => Math.Abs((r.DataTaken.Value - adjustedDate).TotalSeconds))
                    .FirstOrDefault();

                // Step 4: create folder named after earliest image
                string folderName = Path.GetFileNameWithoutExtension(earliestFile);
                string targetFolder = Path.Combine(imagePath, folderName);
                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                // Step 5: find all images within +3s window
                DateTime windowEnd = earliestDate.AddSeconds(3);
                var imagesInWindow = allFiles
                    .Where(f =>
                    {
                        var dt = File.GetLastWriteTime(f);
                        return dt >= earliestDate && dt <= windowEnd;
                    })
                    .ToList();

                // Step 6: move them & write GPS if Excel matched
                foreach (var file in imagesInWindow)
                {
                    string destFile = Path.Combine(targetFolder, Path.GetFileName(file));
                    try
                    {
                        if (File.Exists(destFile)) File.Delete(destFile);
                        File.Move(file, destFile);

                        if (match != null &&
                            double.TryParse(match.Latitude, NumberStyles.Any, CultureInfo.InvariantCulture, out double lat) &&
                            double.TryParse(match.Longitude, NumberStyles.Any, CultureInfo.InvariantCulture, out double lon))
                        {
                            WriteCoordinates(destFile, lat, lon);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to move/write GPS for {file}: {ex.Message}");
                    }
                }

                // Loop continues, new earliest will be picked from remaining files
            }
        }


        public static void WriteCoordinates(string imagePath, double latitude, double longitude)
        {
            try
            {
                string latRef = latitude >= 0 ? "N" : "S";
                string lonRef = longitude >= 0 ? "E" : "W";

                latitude = Math.Abs(latitude);
                longitude = Math.Abs(longitude);

                string exifToolPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Tools\ExifTool\exiftool.exe");
                string args = $"-GPSLatitude={latitude.ToString(CultureInfo.InvariantCulture)} " +
                              $"-GPSLatitudeRef={latRef} " +
                              $"-GPSLongitude={longitude.ToString(CultureInfo.InvariantCulture)} " +
                              $"-GPSLongitudeRef={lonRef} " +
                              $"-overwrite_original \"{imagePath}\"";

                var psi = new ProcessStartInfo
                {
                    FileName = exifToolPath,
                    Arguments = args,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var proc = Process.Start(psi))
                {
                    string output = proc.StandardOutput.ReadToEnd();
                    string error = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();

                    if (!string.IsNullOrEmpty(error))
                        Console.WriteLine("ExifTool error: " + error);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing GPS: {ex.Message}");
            }
        }
    }

    public class AeroRecord
    {
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public DateTime? DataTaken { get; set; }
    }
}
