using ClosedXML.Excel;
using GeographicDynamic_DAL.Interface;
using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.Repository
{
    public class AeroGadagebaWithoutFoldersRepository : IAeroWithoutFolders
    {

        public Result<List<AeroRecordWithoutFolders>> ExcelisWakiTxvaAeroWithoutFolders()
        {
            List<AeroRecordWithoutFolders> ExcelInfo = new List<AeroRecordWithoutFolders>();

            try
            {
                string ExcelPath = @"\\giswebserver\HotSpo - -Pictures\2025_09_23\AleksTest\\GPS_pnt_time.xlsx";
                string ImagePath = @"\\giswebserver\HotSpo - -Pictures\2025_09_23\AleksTest\images";

                // Read Excel
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

                // Write GPS to each image
                WriteGpsFromExcelMatch(ExcelInfo, ImagePath);

                return new Result<List<AeroRecordWithoutFolders>>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                };
            }
            catch (Exception ex)
            {
                return new Result<List<AeroRecordWithoutFolders>>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "ექსელის წაკითხვა ვერ მოხერხდა: " + ex.Message
                };
            }
        }

        private void WriteGpsFromExcelMatch(List<AeroRecordWithoutFolders> excelData, string imagePath)
        {
            var allFiles = Directory.GetFiles(imagePath, "*.*", SearchOption.TopDirectoryOnly).ToList();

            foreach (var file in allFiles)
            {
                try
                {
                    // Step 1: get adjusted image time (-4 hours)
                    DateTime adjustedDate = File.GetLastWriteTime(file).AddHours(-4);

                    // Step 2: find closest Excel time
                    var match = excelData
                        .Where(r => r.DataTaken.HasValue)
                        .OrderBy(r => Math.Abs((r.DataTaken.Value - adjustedDate).TotalSeconds))
                        .FirstOrDefault();

                    if (match != null &&
                        double.TryParse(match.Latitude, NumberStyles.Any, CultureInfo.InvariantCulture, out double lat) &&
                        double.TryParse(match.Longitude, NumberStyles.Any, CultureInfo.InvariantCulture, out double lon))
                    {
                        WriteCoordinates(file, lat, lon);
                        Console.WriteLine($"GPS written to {Path.GetFileName(file)} ({lat}, {lon})");
                    }
                    else
                    {
                        Console.WriteLine($"No valid Excel match for {Path.GetFileName(file)}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process {file}: {ex.Message}");
                }
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

public class AeroRecordWithoutFolders
{
    public string ObjectID { get; set; }
    public DateTime? DataTaken { get; set; }
    public string X { get; set; }
    public string Y { get; set; }
    public string Latitude { get; set; }
    public string Longitude { get; set; }
}
