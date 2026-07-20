using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Ajax.Utilities;
using Swashbuckle.Swagger;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.Http.Results;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/XLSXtoCSVConverterForVAMS")]
    public class XLSXtoCSVConverterForVAMSController : ApiController
    {
        #region Validate the input uploaded file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Input file validated successfully.");
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Convert XLSX to CSV input uploaded file
        [HttpPost]
        [Route("ConvertXLSXtoCSVForVAMS")]
        public HttpResponseMessage ConvertXLSXtoCSVForVAMS(string UploadedInputFileName, string InputFileName, string Delimiter,string inputHospitalName,string ItemOrPO = "Item", bool IsToTrimCellValues = true, bool IsToWriteEachHospitalDataInSeparateFile = true)
        {
            string inputXLSXFilepath = HttpContext.Current.Server.MapPath(@"~/temp/" + InputFileName);
            string tempFolder = Path.GetTempPath();
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(UploadedInputFileName);

            string extension = ".csv";
            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmssffffff");
            string zipFileName = $"ConvertedFiles_{fileNameWithoutExtension}_{timeStamp}.zip";
            string zipPath = Path.Combine(tempFolder, zipFileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));

                Workbook workbook = new Workbook();
                workbook.Open(inputXLSXFilepath);
                var sheet = workbook.Worksheets[0];
                int totalRows = sheet.Cells.MaxDataRow;
                int totalColumns = sheet.Cells.MaxDataColumn;
                int hospitalIDColumnIndex = 0, hospitalNameColumnIndex = 0;
                string safeName = string.Empty;

                List<string> filesToZip = new List<string>();

                // --- Prepare Header String ---
                string[] headerValues = new string[totalColumns + 1];
                for (int c = 0; c <= totalColumns; c++)
                {
                    object cellValue = sheet.Cells[0, c].Value;
                    headerValues[c] = (IsToTrimCellValues ? cellValue?.ToString().Trim() : cellValue?.ToString()) ?? "";

                    if (ItemOrPO.Trim().ToLower() == "po")
                    {
                        if (inputHospitalName.Trim().ToLower() == "uhs")
                        {
                            if (headerValues[c].Equals("Organization", StringComparison.OrdinalIgnoreCase))
                                hospitalIDColumnIndex = c;
                        }
                        else
                        {
                            if (headerValues[c].Equals("Hospital_ID", StringComparison.OrdinalIgnoreCase) ||
                                headerValues[c].Equals("HospitalID", StringComparison.OrdinalIgnoreCase) ||
                                headerValues[c].Equals("Hospital ID", StringComparison.OrdinalIgnoreCase))
                                hospitalIDColumnIndex = c;
                        }

                        if (inputHospitalName.Trim().ToLower() == "centracare")
                        {
                            if (headerValues[c].Equals("Organization", StringComparison.OrdinalIgnoreCase))
                                hospitalNameColumnIndex = c;
                        }
                        else
                        {
                            if (headerValues[c].Equals("Hospital Name", StringComparison.OrdinalIgnoreCase) ||
                                headerValues[c].Equals("HospitalName", StringComparison.OrdinalIgnoreCase))
                                hospitalNameColumnIndex = c;
                        }
                    }
                }
                
                string headerLine = string.Join(Delimiter, headerValues);

                if (IsToWriteEachHospitalDataInSeparateFile)
                {
                    var writers = new Dictionary<string, StreamWriter>();

                    try
                    {
                        for (int r = 1; r <= totalRows; r++)
                        {
                            string[] rowData = new string[totalColumns + 1];
                            for (int c = 0; c <= totalColumns; c++)
                            {
                                object cellValue = sheet.Cells[r, c].Value;
                                rowData[c] = (IsToTrimCellValues ? cellValue?.ToString().Trim() : cellValue?.ToString()) ?? "";
                            }

                            string hospitalName = rowData[hospitalNameColumnIndex];
                            string hospitalID = rowData[hospitalIDColumnIndex];

                            if (inputHospitalName.Trim().ToLower().Contains("albany"))
                            {
                                if (ItemOrPO.Trim().ToLower() == "item")
                                    safeName = "Albany_Item";
                                else if (ItemOrPO.Trim().ToLower() == "po")
                                {
                                    if(hospitalName.Trim().ToLower()=="amh")
                                        safeName = "A_Albany Medical Center_PO";
                                    else if (hospitalName.Trim().ToLower() == "cmh")
                                        safeName = "A_Columbia Memorial Hospital_PO";
                                    else if (hospitalName.Trim().ToLower() == "gfh")
                                        safeName = "A_Glens Falls Hospital_PO";
                                    else if (hospitalName.Trim().ToLower() == "tsh")
                                        safeName = "A_Saratoga Hospital_PO";
                                    else
                                        safeName = "A_" + string.Join("_", hospitalName.Split(Path.GetInvalidFileNameChars())) + "_PO";
                                }
                            }
                            else if (inputHospitalName.Trim().ToLower().Contains("mercy"))
                            {
                                    if (ItemOrPO.Trim().ToLower() == "item")
                                        safeName = "Mercy_Item";
                                    else if (ItemOrPO.Trim().ToLower() == "po")
                                        safeName = "Mercy_PO";
                            }
                            else if (inputHospitalName.Trim().ToLower().Contains("minnesota"))
                            {
                                if (ItemOrPO.Trim().ToLower() == "item")
                                    safeName = "Minnesota_Item";
                                else if (ItemOrPO.Trim().ToLower() == "po")
                                    safeName = "Minnesota_PO";
                            }
                            else if (inputHospitalName.Trim().ToLower().Contains("texas"))
                            {
                                if (ItemOrPO.Trim().ToLower() == "item")
                                    safeName = "Texas_Item";
                                else if (ItemOrPO.Trim().ToLower() == "po")
                                    safeName = "Texas_PO";
                            }
                            else if (inputHospitalName.Trim().ToLower().Contains("childrensconnecticut"))
                            {
                                if (ItemOrPO.Trim().ToLower() == "item")
                                    safeName = "ChildrensConnecticut_Item";
                                else if (ItemOrPO.Trim().ToLower() == "po")
                                    safeName = "ChildrensConnecticut_PO";
                            }
                            else if (inputHospitalName.Trim().ToLower().Contains("nuvance"))
                            {
                                if (ItemOrPO.Trim().ToLower() == "item")
                                    safeName = "Nuvance_Item";
                                else if (ItemOrPO.Trim().ToLower() == "po")
                                    safeName = "N_" + string.Join("_", hospitalName.Split(Path.GetInvalidFileNameChars())) + "_PO";
                            }
                            else if (inputHospitalName.Trim().ToLower().Contains("uhs"))
                            {
                                if (ItemOrPO.Trim().ToLower() == "item")
                                    safeName = "UHS_Item";
                                else if (ItemOrPO.Trim().ToLower() == "po")
                                {
                                    if (hospitalName.Trim().ToLower() == "uhs corporate")
                                        safeName = "01_UHS CORPORATE_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs hospitals")
                                        safeName = "02_UHS HOSPITALS_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs chenango memorial hospital")
                                        safeName = "03_UHS CHENANGO MEMORIAL HOSPITAL_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs medical group")
                                        safeName = "04_UHS MEDICAL GROUP_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs senior living at ideal")
                                        safeName = "05_UHS SENIOR LIVING AT IDEAL_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs senior living at ideal hc")
                                        safeName = "06_UHS SENIOR LIVING AT IDEAL HC_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs home care tthh")
                                        safeName = "07_UHS HOME CARE TTHH_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs home care phc")
                                        safeName = "08_UHS HOME CARE PHC_PO";
                                    else if (hospitalName.Trim().ToLower() == "uhs delaware valley hospital")
                                        safeName = "09_UHS DELAWARE VALLEY HOSPITAL_PO";
                                    else
                                        safeName = string.Join("_", hospitalName.Split(Path.GetInvalidFileNameChars())) + "_PO";
                                }
                                // safeName = hospitalID + " " + string.Join("_", hospitalName.Split(Path.GetInvalidFileNameChars())) + "_PO";
                            }
                            else if (inputHospitalName.Trim().ToLower().Contains("centracare"))
                            {
                                if (ItemOrPO.Trim().ToLower() == "item")
                                    safeName = "CentraCare_Item";
                                else if (ItemOrPO.Trim().ToLower() == "po")
                                    safeName = "C_" + string.Join("_", hospitalName.Split(Path.GetInvalidFileNameChars())) + "_PO";
                            }

                            string hospitalFilePath = Path.Combine(tempFolder, $"{safeName}{extension}");

                            if (!writers.ContainsKey(safeName))
                            {
                                StreamWriter sw = new StreamWriter(hospitalFilePath, false);
                                sw.WriteLine(headerLine);

                                writers.Add(safeName, sw);
                                filesToZip.Add(hospitalFilePath);
                            }
                            writers[safeName].WriteLine(string.Join(Delimiter, rowData));
                        }

                        // --- CRITICAL CHANGE START ---
                        // 1. Close and Dispose writers immediately after the loop
                        foreach (var writer in writers.Values)
                        {
                            writer.Dispose();
                        }
                        writers.Clear(); // Clear the dictionary so 'finally' doesn't try to dispose again
                        // --- CRITICAL CHANGE END ---

                        // 2. Create the ZIP file
                        using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                        {
                            foreach (string filePath in filesToZip)
                            {
                                if (File.Exists(filePath))
                                    zip.AddFile(filePath, "");
                            }
                            zip.Save(zipPath);
                        }

                        // 3. Cleanup
                        foreach (string filePath in filesToZip)
                        {
                            if (File.Exists(filePath)) File.Delete(filePath);
                        }

                        if (File.Exists(inputXLSXFilepath)) File.Delete(inputXLSXFilepath);
                    }
                    finally
                    {
                        foreach (var writer in writers.Values) { writer.Dispose(); }
                    }


                    if (File.Exists(inputXLSXFilepath)) File.Delete(inputXLSXFilepath);
                    return Request.CreateResponse(HttpStatusCode.OK, zipPath);
                }
                else
                {
                    if (inputHospitalName.Trim().ToLower().Contains("albany"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "Albany_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "Albany_PO";
                    }
                    else if (inputHospitalName.Trim().ToLower().Contains("mercy"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "Mercy_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "Mercy_PO";
                    }
                    else if (inputHospitalName.Trim().ToLower().Contains("minnesota"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "Minnesota_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "Minnesota_PO";
                    }
                    else if (inputHospitalName.Trim().ToLower().Contains("texas"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "Texas_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "Texas_PO";
                    }
                    else if (inputHospitalName.Trim().ToLower().Contains("childrensconnecticut"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "ChildrensConnecticut_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "ChildrensConnecticut_PO";
                    }
                    else if (inputHospitalName.Trim().ToLower().Contains("nuvance"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "Nuvance_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "Nuvance_PO";
                    }
                    else if (inputHospitalName.Trim().ToLower().Contains("uhs"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "UHS_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "UHS_PO";
                    }
                    else if (inputHospitalName.Trim().ToLower().Contains("centracare"))
                    {
                        if (ItemOrPO.Trim().ToLower() == "item")
                            safeName = "CentraCare_Item";
                        else if (ItemOrPO.Trim().ToLower() == "po")
                            safeName = "CentraCare_PO";
                    }

                    string hospitalFilePath = Path.Combine(tempFolder, $"{safeName}{extension}");
                    using (StreamWriter sw = new StreamWriter(hospitalFilePath, false))
                    {
                        // Write Header once at the top
                        sw.WriteLine(headerLine);

                        // Process data starting from row 1
                        for (int r = 1; r <= totalRows; r++)
                        {
                            string[] rowData = new string[totalColumns + 1];
                            for (int c = 0; c <= totalColumns; c++)
                            {
                                object cellValue = sheet.Cells[r, c].Value;
                                rowData[c] = (IsToTrimCellValues ? cellValue?.ToString().Trim() : cellValue?.ToString()) ?? "";
                            }
                            sw.WriteLine(string.Join(Delimiter, rowData));
                        }
                    }

                    if (File.Exists(inputXLSXFilepath)) File.Delete(inputXLSXFilepath);
                    return Request.CreateResponse(HttpStatusCode.OK, hospitalFilePath);
                }
            }
            catch (Exception ex)
            {
                if (File.Exists(inputXLSXFilepath)) File.Delete(inputXLSXFilepath);
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion
    }
}
