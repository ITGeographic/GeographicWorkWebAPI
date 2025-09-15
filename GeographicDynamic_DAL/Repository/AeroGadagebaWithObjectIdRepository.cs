using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using GeographicDynamic_DAL.Interface;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using Font = System.Drawing.Font;
using System.Drawing.Imaging;
using System.Runtime.Serialization;
using RedCorners.ExifLibrary;
using System.Diagnostics;
using System.Globalization;
using static System.Net.Mime.MediaTypeNames;

namespace GeographicDynamic_DAL.Repository
{
    public class AeroGadagebaWithObjectIdRepository : IAeroWithObjectId
    {
        public Result<List<AeroRecordWithObjectId>> ExcelisWakiTxvaAeroWithObjectId()
        {
            List<AeroRecordWithObjectId> ExcelInfo = new List<AeroRecordWithObjectId>();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkbook = null;
            _Worksheet xlWorksheet = null;
            Microsoft.Office.Interop.Excel.Range xlRange = null;

            try
            {
                string ExcelPath = @"D:\Projects\2025\QarsaffariDatvlebi\TestAero\aero.xlsx";
                xlWorkbook = xlApp.Workbooks.Open(ExcelPath);
                xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;

                for (int i = 2; i <= rowCount; i++) // assuming first row is headers
                {
                    var objectID = xlRange.Cells[i, 1]?.Value2?.ToString();
                    var cellDate = xlRange.Cells[i, 2]?.Value2;
                    DateTime? dataTaken = null;

                    if (cellDate != null)
                    {
                        if (cellDate is double d)
                            dataTaken = DateTime.FromOADate(d);
                        else if (cellDate is string s && DateTime.TryParse(s, out var parsed))
                            dataTaken = parsed;
                    }

                    var xValue = xlRange.Cells[i, 3]?.Value2?.ToString();
                    var yValue = xlRange.Cells[i, 4]?.Value2?.ToString();

                    ExcelInfo.Add(new AeroRecordWithObjectId
                    {
                        ObjectID = objectID,
                        DataTaken = dataTaken,
                        X = xValue,
                        Y = yValue
                    });
                }

                // ===== Organize images & write coordinates =====
                OrganizeImagesAndWriteCoordinates(ExcelInfo, @"D:\Projects\2025\QarsaffariDatvlebi\TestAero\images");

                return new Result<List<AeroRecordWithObjectId>>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                };
            }
            catch (Exception ex)
            {
                return new Result<List<AeroRecordWithObjectId>>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "ექსელის წაკითხვა ვერ მოხერხდა: " + ex.Message
                };
            }
            finally
            {
                if (xlRange != null) Marshal.FinalReleaseComObject(xlRange);
                if (xlWorksheet != null) Marshal.FinalReleaseComObject(xlWorksheet);
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    Marshal.FinalReleaseComObject(xlWorkbook);
                }
                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
            }
        }

        private void OrganizeImagesAndWriteCoordinates(List<AeroRecordWithObjectId> excelData, string imagePath)
        {
            // Step 1: Preload all image dates to reduce repeated file access
            var allFiles = Directory.GetFiles(imagePath, "*.*", SearchOption.AllDirectories);
            var fileDates = new Dictionary<string, DateTime>();

            foreach (var file in allFiles)
            {
                try
                {
                    DateTime? photoDate = GetPhotoDateTaken(file);
                    if (photoDate.HasValue)
                        fileDates[file] = photoDate.Value;
                }
                catch
                {
                    // skip unreadable images
                }
            }
            // Step 2: Move images into ObjectID folders using time ranges
            for (int i = 0; i < excelData.Count; i++)
            {
                var currentRecord = excelData[i];
                if (currentRecord.DataTaken == null) continue;

                DateTime startTime = currentRecord.DataTaken.Value;
                DateTime? endTime = (i < excelData.Count - 1)
                                    ? excelData[i + 1].DataTaken
                                    : null; // no upper limit for the last one

                // Always create folder for current ObjectID
                string targetFolder = Path.Combine(imagePath, currentRecord.ObjectID);
                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                foreach (var kvp in fileDates)
                {
                    string file = kvp.Key;
                    DateTime photoDate = kvp.Value;

                    bool inRange = endTime.HasValue
                        ? (photoDate >= startTime && photoDate < endTime.Value)
                        : (photoDate >= startTime); // last range gets all remaining images

                    if (inRange)
                    {
                        string destFile = Path.Combine(targetFolder, Path.GetFileName(file));
                        try
                        {
                            if (File.Exists(destFile)) File.Delete(destFile);
                            File.Move(file, destFile);
                        }
                        catch
                        {
                            // ignore move failures
                        }
                    }
                }
            }


            // Step 3: Write coordinates on images
            var objectFolders = Directory.GetDirectories(imagePath);
            foreach (var folder in objectFolders)
            {
                string folderName = Path.GetFileName(folder);
                foreach (var imageFile in Directory.GetFiles(folder))
                {
                    try
                    {
                        var photoDate = GetPhotoDateTaken(imageFile);
                        if (!photoDate.HasValue) continue;

                        var match = excelData.FirstOrDefault(r =>
                            r.ObjectID == folderName &&
                            r.DataTaken.HasValue &&
                            Math.Abs((r.DataTaken.Value - photoDate.Value).TotalMinutes) <= 5);

                        if (match != null)
                        {
                            double x = Convert.ToDouble(match.X, CultureInfo.InvariantCulture);
                            double y = Convert.ToDouble(match.Y, CultureInfo.InvariantCulture);
                            WriteCoordinates(imageFile, x, y);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to write GPS for {imageFile}: {ex.Message}");
                    }
                }
            }
        }


        private DateTime? GetPhotoDateTaken(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var img = System.Drawing.Image.FromStream(fs))
                {
                    const int PropertyTagDateTaken = 36867;

                    if (img.PropertyIdList.Contains(PropertyTagDateTaken))
                    {
                        var prop = img.GetPropertyItem(PropertyTagDateTaken);
                        string dateStr = System.Text.Encoding.ASCII.GetString(prop.Value).Trim('\0');
                        if (DateTime.TryParseExact(dateStr, "yyyy:MM:dd HH:mm:ss",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out DateTime dt))
                        {
                            return dt;
                        }
                    }
                }
            }
            catch { }

            // fallback
            return File.GetLastWriteTime(path);
        }
        // Helper method
        static void SetProperty(int id, double value)
        {
            // კონვერტირება დეგრადებიდან DMS-ში (degrees/minutes/seconds)
            uint[] rational = ConvertToRational(value);

        }
        static uint[] ConvertToRational(double decimalDegrees)
        {
            decimalDegrees = Math.Abs(decimalDegrees);
            int degrees = (int)decimalDegrees;
            double minutesFloat = (decimalDegrees - degrees) * 60;
            int minutes = (int)minutesFloat;
            double seconds = (minutesFloat - minutes) * 60;

            return new uint[]
            {
            (uint)degrees, 1,
            (uint)minutes, 1,
            (uint)(seconds * 100), 100 // Fractional seconds
            };
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


        // --- Helper: Convert UTM to Lat/Lon ---
        private static (double Latitude, double Longitude) UtmToLatLon(double easting, double northing, int zoneNumber, bool isNorthernHemisphere)
        {
            const double a = 6378137.0; // WGS84 major axis
            const double e = 0.0818191908; // WGS84 eccentricity
            const double k0 = 0.9996;

            double x = easting - 500000.0;
            double y = northing;
            if (!isNorthernHemisphere)
                y -= 10000000.0;

            double lonOrigin = (zoneNumber - 1) * 6 - 180 + 3;

            double eccPrimeSquared = (e * e) / (1 - e * e);
            double M = y / k0;
            double mu = M / (a * (1 - e * e / 4.0 - 3 * e * e * e * e / 64.0 - 5 * e * e * e * e * e * e / 256.0));

            double e1 = (1 - Math.Sqrt(1 - e * e)) / (1 + Math.Sqrt(1 - e * e));

            double J1 = 3 * e1 / 2 - 27 * e1 * e1 * e1 / 32.0;
            double J2 = 21 * e1 * e1 / 16 - 55 * e1 * e1 * e1 * e1 / 32.0;
            double J3 = 151 * e1 * e1 * e1 / 96.0;
            double J4 = 1097 * e1 * e1 * e1 * e1 / 512.0;

            double fp = mu + J1 * Math.Sin(2 * mu)
                          + J2 * Math.Sin(4 * mu)
                          + J3 * Math.Sin(6 * mu)
                          + J4 * Math.Sin(8 * mu);

            double C1 = eccPrimeSquared * Math.Cos(fp) * Math.Cos(fp);
            double T1 = Math.Tan(fp) * Math.Tan(fp);
            double N1 = a / Math.Sqrt(1 - e * e * Math.Sin(fp) * Math.Sin(fp));
            double R1 = a * (1 - e * e) / Math.Pow(1 - e * e * Math.Sin(fp) * Math.Sin(fp), 1.5);
            double D = x / (N1 * k0);

            double Q1 = N1 * Math.Tan(fp) / R1;
            double Q2 = (D * D / 2.0);
            double Q3 = (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * eccPrimeSquared) * D * D * D * D / 24.0;
            double Q4 = (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * eccPrimeSquared - 3 * C1 * C1) * D * D * D * D * D * D / 720.0;
            double lat = fp - Q1 * (Q2 - Q3 + Q4);

            double Q5 = D;
            double Q6 = (1 + 2 * T1 + C1) * D * D * D / 6;
            double Q7 = (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * eccPrimeSquared + 24 * T1 * T1)
                          * D * D * D * D * D / 120.0;
            double lon = lonOrigin + (Q5 - Q6 + Q7) / Math.Cos(fp);

            lat = lat * 180.0 / Math.PI;
            lon = lon * 180.0 / Math.PI;

            return (lat, lon);
        }


    }

    public class AeroRecordWithObjectId
    {
        public string ObjectID { get; set; }
        public DateTime? DataTaken { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
    }
}