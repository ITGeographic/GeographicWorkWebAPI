using ClosedXML.Excel;
using GeographicDynamic_DAL.Interface;
using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using MetadataExtractor;
using MetadataExtractor.Formats.Exif;
using System.Net;

namespace GeographicDynamic_DAL.Repository
{
    public class AeroGadagebaWithoutFolderingSecondApproach : IAeroWithoutFolders
    {

        public Result<List<AeroRecordWithoutFolders>> ExcelisWakiTxvaAeroWithoutFolders()
        {
            List<AeroRecordWithoutFolders> ExcelInfo = new List<AeroRecordWithoutFolders>();
            try
            {
                string ExcelPath = @"\\giswebserver\HotSpo Pictures\2025_09_24\AleksTest\GPS_pnt_time.xlsx";
                string ImagePath = @"\\giswebserver\HotSpo Pictures\2025_09_24\AleksTest\images";
                string ResultPath = Path.Combine(ImagePath, "Result");
                System.IO.Directory.CreateDirectory(ResultPath);

                // 1. Read Excel
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

                        ExcelInfo.Add(new AeroRecordWithoutFolders
                        {
                            Latitude = latStr,
                            Longitude = lonStr,
                            DataTaken = dataTaken
                        });
                    }
                }

                // 2. Iterate images and match
                var allFiles = System.IO.Directory.GetFiles(ImagePath, "*.jpg", SearchOption.TopDirectoryOnly);

                foreach (var file in allFiles)
                {
                    try
                    {
                        DateTime? photoTime = GetPhotoDateTaken(file);
                        if (photoTime == null) continue;

                        // Adjust (+4h)
                        DateTime adjustedDate = photoTime.Value.AddHours(-4);

                        // Find closest Excel point within 3 seconds
                        var match = ExcelInfo
                            .Where(r => r.DataTaken.HasValue)
                            .OrderBy(r => Math.Abs((r.DataTaken.Value - adjustedDate).TotalSeconds))
                            .FirstOrDefault();

                        if (match != null &&
                            Math.Abs((match.DataTaken.Value - adjustedDate).TotalSeconds) <= 3 &&
                            double.TryParse(match.Latitude, NumberStyles.Any, CultureInfo.InvariantCulture, out double lat) &&
                            double.TryParse(match.Longitude, NumberStyles.Any, CultureInfo.InvariantCulture, out double lon))
                        {
                            // Write GPS EXIF
                            WriteCoordinates(file, lat, lon);

                            // Copy to Result folder
                            string destFile = Path.Combine(ResultPath, Path.GetFileName(file));
                            File.Copy(file, destFile, true);

                            Console.WriteLine($"[{Path.GetFileName(file)}] matched → GPS ({lat}, {lon}) written and copied.");
                        }
                        else
                        {
                            string withoutCoordinates = @"\\giswebserver\HotSpo Pictures\2025_09_23\AleksTest\images\withoutCoordinates";



                            string destFileWithoutCoordinates = Path.Combine(withoutCoordinates, Path.GetFileName(file));
                            Console.WriteLine($"[{Path.GetFileName(file)}] no Excel match (diff > 3s).");
                            File.Copy(file, destFileWithoutCoordinates, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process {file}: {ex.Message}");
                    }
                }

                return new Result<List<AeroRecordWithoutFolders>>
                {
                    Success = true,
                    StatusCode = HttpStatusCode.OK,
                    Value = ExcelInfo //
                };
            }
            catch (Exception ex)
            {
                return new Result<List<AeroRecordWithoutFolders>>
                {
                    Success = false,
                    StatusCode = HttpStatusCode.BadGateway,
                    Message = "ექსელის წაკითხვა ვერ მოხერხდა: " + ex.Message,
                    Value = new List<AeroRecordWithoutFolders>() // ✅ empty list
                };
            }
        }

        private DateTime? GetPhotoDateTaken(string path)
        {
            try
            {
                var directories = ImageMetadataReader.ReadMetadata(path);
                var subIfd = directories.OfType<ExifSubIfdDirectory>().FirstOrDefault();
                return subIfd?.GetDateTime(ExifDirectoryBase.TagDateTimeOriginal);
            }
            catch
            {
                return null;
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


}
