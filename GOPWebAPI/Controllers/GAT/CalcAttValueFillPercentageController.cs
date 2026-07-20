using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;


namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/CalcAttValueFillPercentage")]
    public class CalcAttValueFillPercentageController : ApiController
    {
        #region Validate the input uploaded file first worksheet
        [HttpPost]
        [Route("ValidateInputFileFirstWorksheet")]
        public HttpResponseMessage ValidateInputFileFirstWorksheet(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIFFW = wbIF.Worksheets[0];      //Input File First Worksheet (Attributes and Values)

                #region Check uploaded file first worksheet has columns as per template
                if (wsIFFW.Cells[0, 0].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells[0, 1].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Noun'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIFFW.Cells[0, 2].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                string hdr;
                int attNo = 1;
                for (int aCol = 1; aCol <= 99; aCol += 2)
                {
                    hdr = wsIFFW.Cells[0, aCol + 2].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCol + 3).ToString() + "] must contain 'Attribute " + attNo.ToString() + "' .Please download format from Template.");
                    }

                    hdr = wsIFFW.Cells[0, aCol + 3].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE VALUE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCol + 4).ToString() + "] must contain 'Attribute Value" + attNo.ToString() + "' .Please download format from Template.");
                    }
                    attNo++;
                }

                if (wsIFFW.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "success");
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

        #region Validate the input uploaded file second worksheet
        [HttpPost]
        [Route("ValidateInputFileSecondWorksheet")]
        public HttpResponseMessage ValidateInputFileSecondWorksheet(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                
                if(wbIF.Worksheets.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file should have 2 worksheets. Please select a valid file.");
                }

                Worksheet wsIFSW = wbIF.Worksheets[1];      //Input File Second Worksheet (Attribute Mandatory / Optional)

                #region Check uploaded file second worksheet has columns as per template
                if (wsIFSW.Cells[0, 0].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Noun'  not found in input file second worksheet. Please select a valid file.");
                }

                if (wsIFSW.Cells[0, 1].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Modifier' not found in input file second worksheet. Please select a valid file.");
                }

                if (wsIFSW.Cells[0, 2].StringValue.Trim().ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Attribute' not found in input file second worksheet. Please select a valid file.");
                }

                if (wsIFSW.Cells[0, 3].StringValue.Trim().ToUpper() != "MANDATORY/OPTIONAL")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'Mandatory/Optional' not found in input file second worksheet. Please select a valid file.");
                }

                if (wsIFSW.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File second Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "success");
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

        #region Calculate Attribute Value Fill Percentage and Write results to output file
        [HttpPost]
        [Route("CalcAttValueFillPercentageAndWriteOutputToExcel")]
        public HttpResponseMessage CalcAttValueFillPercentageAndWriteOutputToExcel(string UploadedInputFileName, string FileName)
        {
            try
            {
                string InputFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

                #region Setting up the workbook
                AsposeHelpers asposeHelpers = new AsposeHelpers();
                Workbook wb = asposeHelpers.GetWorkbook();
                wb.LoadData(InputFilePath);
                #endregion

                Worksheet wsIAV = wb.Worksheets[0];
                int maxIRows = wsIAV.Cells.MaxRow;
                int maxIColumns = wsIAV.Cells.MaxColumn;

                Worksheet wsMOA = wb.Worksheets[1];
                int maxMOARows = wsMOA.Cells.MaxRow;

                #region Adding Input Mandatory Attributes To List (instead of creating another class, using existing)
                List<ProjectNounModifierAttribute> MandatoryAttributesList = new List<ProjectNounModifierAttribute>();
                for (int mRow = 1; mRow <= maxMOARows; mRow++)
                {
                    if (wsMOA.Cells[mRow, 3].StringValue.Trim().ToUpper() == "M" || wsMOA.Cells[mRow, 3].StringValue.Trim().ToUpper() == "MANDATORY")
                    {
                        MandatoryAttributesList.Add(new ProjectNounModifierAttribute()
                        {
                            Noun = wsMOA.Cells[mRow, 0].StringValue.Trim(),
                            Modifier = wsMOA.Cells[mRow, 1].StringValue.Trim(),
                            AttributeName = wsMOA.Cells[mRow, 2].StringValue.Trim()
                        });
                    }
                }
                #endregion

                #region Inserting result columns
                wsIAV.Cells.InsertColumn(0);
                wsIAV.Cells.InsertColumn(0);
                wsIAV.Cells.InsertColumn(0);
                wsIAV.Cells.InsertColumn(0);
                wsIAV.Cells.InsertColumn(0);
                wsIAV.Cells.InsertColumn(0);

                wsIAV.Cells[0, 0].PutValue("No. of Attributes");
                wsIAV.Cells[0, 1].PutValue("No. of Values");
                wsIAV.Cells[0, 2].PutValue("Att. Value Fill % for all attributes");
                wsIAV.Cells[0, 3].PutValue("No. of Mandatory Attributes");
                wsIAV.Cells[0, 4].PutValue("No. of Mandatory Values");
                wsIAV.Cells[0, 5].PutValue("Att. Value Fill % for Mandatory attributes");
                #endregion

                int NoOfAttributesCounter = 0, NoOfValuesCounter = 0, NoOfMandatoryAttributesCounter = 0, NoOfMandatoryValuesCounter = 0;
                decimal AttributeValueFillPercentage = 0, MandatoryValuesFillPercentage = 0;
                string Noun = string.Empty, Modifier = string.Empty, Attribute = string.Empty, AttributeValue = string.Empty;

                #region Processing Rows starts
                for (int row = 1; row <= maxIRows; row++)
                {
                    #region Calculate No. of Attributes and No. of Values
                    NoOfAttributesCounter = 0;
                    NoOfValuesCounter = 0;
                    for (int col = 9; col <= 108; col += 2)
                    {
                        if (!string.IsNullOrEmpty(wsIAV.Cells[row, col].StringValue.Trim()))
                            NoOfAttributesCounter++;
                        else
                            break;

                        if (!string.IsNullOrEmpty(wsIAV.Cells[row, col + 1].StringValue.Trim()))
                            NoOfValuesCounter++;

                    }
                    wsIAV.Cells[row, 0].PutValue(NoOfAttributesCounter);
                    wsIAV.Cells[row, 1].PutValue(NoOfValuesCounter);
                    #endregion

                    #region Calculate Att. Value Fill Percentage
                    AttributeValueFillPercentage = 0;
                    if (NoOfAttributesCounter > 0 && NoOfValuesCounter > 0)
                        AttributeValueFillPercentage = Convert.ToDecimal(string.Format("{0:0.00}", (NoOfValuesCounter * 100.00) / (NoOfAttributesCounter * 1.00)));

                    wsIAV.Cells[row, 2].PutValue(AttributeValueFillPercentage);
                    #endregion

                    #region No. of Mandatory Attributes, No. of Mandatory Values, Mandatory Values Fill Percentage
                    NoOfMandatoryAttributesCounter = 0; NoOfMandatoryValuesCounter = 0; MandatoryValuesFillPercentage = 0;
                    Noun = wsIAV.Cells[row, 7].StringValue.Trim();
                    Modifier = wsIAV.Cells[row, 8].StringValue.Trim();
                    if (!string.IsNullOrEmpty(Noun) && !string.IsNullOrEmpty(Modifier))
                    {
                        for (int col = 9; col <= 108; col += 2)
                        {
                            Attribute = wsIAV.Cells[row, col].StringValue.Trim();

                            if (string.IsNullOrEmpty(Attribute))
                                break;

                            AttributeValue = wsIAV.Cells[row, col + 1].StringValue.Trim();
                            if (MandatoryAttributesList
                                .Where(ml => ml.Noun.ToUpper() == Noun.ToUpper() &&
                                             ml.Modifier.ToUpper() == Modifier.ToUpper() &&
                                             ml.AttributeName.ToUpper() == Attribute.ToUpper()).Count() > 0)
                            {
                                NoOfMandatoryAttributesCounter++;
                                if (!string.IsNullOrEmpty(AttributeValue))
                                    NoOfMandatoryValuesCounter++;
                            }
                        }
                        wsIAV.Cells[row, 3].PutValue(NoOfMandatoryAttributesCounter);
                        wsIAV.Cells[row, 4].PutValue(NoOfMandatoryValuesCounter);
                        if (NoOfMandatoryAttributesCounter > 0 && NoOfMandatoryValuesCounter > 0)
                            MandatoryValuesFillPercentage = Convert.ToDecimal(string.Format("{0:0.00}", (NoOfMandatoryValuesCounter * 100.00) / (NoOfMandatoryAttributesCounter * 1.00)));
                        wsIAV.Cells[row, 5].PutValue(MandatoryValuesFillPercentage);
                    }
                    #endregion
                }
                #endregion

                if (File.Exists(InputFilePath))
                    File.Delete(InputFilePath);

                wsIAV.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wb.Save(filename);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion
    }
}
