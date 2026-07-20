using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using Microsoft.Ajax.Utilities;
using Swashbuckle.Swagger;
using System;
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
    [RoutePrefix("api/XLSXtoTXT")]
    public class XLSXtoTXTController : ApiController
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

        #region Convert XLSX to TXT input uploaded file
        [HttpPost]
        [Route("ConvertXLSXtoTXT")]
        public HttpResponseMessage ConvertXLSXtoTXT(string UploadedInputFileName, string InputFileName, string Delimiter, bool IsToTrimCellValues = true)
        {
            // 1. Define paths correctly
            string inputXLSXFilepath = HttpContext.Current.Server.MapPath(@"~/temp/" + InputFileName);

            // Use Path.Combine to save to the system's actual Temp folder
            string tempFolder = Path.GetTempPath();
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(UploadedInputFileName);
            string outputFileName = "Output_" + fileNameWithoutExtension + ".txt";
            string outputTXTFilepath = Path.Combine(tempFolder, outputFileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                // Load Aspose License
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));

                // Initialize Workbook (Modern Aspose uses the constructor to open files)
                Workbook workbook = new Workbook();
                workbook.Open(inputXLSXFilepath);
                var sheet = workbook.Worksheets[0];

                int totalRows = sheet.Cells.MaxDataRow;
                int totalColumns = sheet.Cells.MaxDataColumn;

                // 2. Open StreamWriter ONCE outside the loop (much faster)
                using (StreamWriter sw = new StreamWriter(outputTXTFilepath, false))
                {
                    for (int r = 0; r <= totalRows; r++) // Starting from 0 to include headers, or 1 for data only
                    {
                        string[] splittedValues = new string[totalColumns + 1];
                        for (int c = 0; c <= totalColumns; c++)
                        {
                            object cellValue = sheet.Cells[r, c].Value;
                            string value = cellValue == null ? "" : cellValue.ToString();

                            if (IsToTrimCellValues)
                                value = value.Trim();

                            splittedValues[c] = value;
                        }

                        string line = string.Join(Delimiter, splittedValues);
                        sw.WriteLine(line);
                    }
                }

                // 3. Return the filename in the response
                return Request.CreateResponse(HttpStatusCode.OK, outputTXTFilepath);
            }
            catch (Exception ex)
            {
                // Cleanup resources on failure
                if (File.Exists(inputXLSXFilepath)) File.Delete(inputXLSXFilepath);
                if (File.Exists(outputTXTFilepath)) File.Delete(outputTXTFilepath);

                ExceptionLogging.SendExceptionToDB(ex);

                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion
    }
}
