using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;
using Swashbuckle.Swagger;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Web.DynamicData;
using System.Web.Http;
using System.Web.Http.Results;

namespace GOPWebAPI.Controllers.GAT
{

    [RoutePrefix("api/DuplicatesCheckOnAttributes")]
    public class DuplicatesCheckOnAttributesController : ApiController
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

                #region Validating the first worksheet
                Worksheet ws1 = wbIF.Worksheets[0];
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 2].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [C1] with Value 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                #region Validating first worksheet attributes columns header
                string hdr;
                int attNo = 1;
                for (int aCtr = 2; aCtr <= 100; aCtr += 2)
                {
                    hdr = ws1.Cells[0, aCtr + 1].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 2).ToString() + "] must contain 'Attribute" + attNo.ToString() + "' in input file first worksheet.Please download format from Template.");
                    }

                    hdr = ws1.Cells[0, aCtr + 2].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE VALUE " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 3).ToString() + "] must contain 'Attribute Value " + attNo.ToString() + "' in input file first worksheet.Please download format from Template.");
                    }
                    attNo++;
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

        #region Validate the attribute list uploaded file
        [HttpPost]
        [Route("ValidateAttributeListFile")]
        public HttpResponseMessage ValidateAttributeListFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);
                
                string ext = Path.GetExtension(FileName).ToLower();
                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid attribute list file of xlsx format only");
                
                Workbook wbALF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbALF.Open(UploadedFilepath);

                #region Validating the first worksheet
                Worksheet ws1 = wbALF.Worksheets[0];
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Noun' not found in attribute list file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Modifier' not found in attribute list file first worksheet. Please select a valid file.");
                }

                int attNo = 1;
                for (int AttColNo = 1; AttColNo <= 50; AttColNo++)
                {
                    string hdr = ws1.Cells[0, AttColNo + 1].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Attribute list file first worksheet Column No. " + (AttColNo + 2).ToString() + " must contain 'Attribute " + attNo.ToString() + "' heading. Please download format from Template.");
                    }

                    attNo++;
                }

                if (ws1.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Attribute List File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Attribute list file validated successfully.");
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

        #region Check for Duplicate Rows and Write the duplicate remarks in Output to file
        [HttpPost]
        [Route("DuplicateCheckOnAttributesAndWriteOutputToExcel")]
        public HttpResponseMessage DuplicateCheckOnAttributesAndWriteOutputToExcel(string InputFileName, string UploadedInputFileName, string AttributeListFileName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
                string UploadedAttributeListFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + AttributeListFileName);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                Workbook wbALF = new Workbook();
                Aspose.Cells.License l1 = new Aspose.Cells.License();
                l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbALF.Open(UploadedAttributeListFilepath);
                Worksheet wsALF = wbALF.Worksheets[0];

                int IFMaxRows = wsIF.Cells.MaxRow;
                int IFMaxColumns = wsIF.Cells.MaxColumn;
                int ALFMaxRows = wsALF.Cells.MaxRow;
                int ALFMaxColumns = wsALF.Cells.MaxColumn;

                int ObservationColNo = IFMaxColumns + 1;
                wsIF.Cells[0, ObservationColNo].PutValue("Observation");

                List<int> IFAVColNoList = new List<int>();
                string iNoun = string.Empty, iModifier = string.Empty, iAttributeName = string.Empty;
                bool iAttributeNameFoundInAttributeList = false, AreAllColumnValuesMatching = false, IsDuplicateCounterUsed = false;
                int ALFNMFoundRowNo = 0, DuplicateSetCounter = 1;

                Aspose.Cells.Style styleLightBlueFillColor = wsIF.Cells[0, ObservationColNo + 1].GetStyle();

                //Setting fill-color style
                styleLightBlueFillColor.ForegroundColor = System.Drawing.Color.FromArgb(0, 138, 199, 219);
                styleLightBlueFillColor.BackgroundColor = System.Drawing.Color.FromArgb(0, 138, 199, 219);
                styleLightBlueFillColor.Pattern = BackgroundType.VerticalStripe;
                styleLightBlueFillColor.Font.Color = System.Drawing.Color.Black;

                #region Processing Input File Rows starts
                for (int iRow = 1; iRow <= IFMaxRows; iRow++)
                {
                    if (string.IsNullOrEmpty(wsIF.Cells[iRow, ObservationColNo].StringValue.Trim()))
                    {
                        iNoun = wsIF.Cells[iRow, 1].StringValue.Trim();
                        iModifier = wsIF.Cells[iRow, 2].StringValue.Trim();

                        if (string.IsNullOrEmpty(iNoun))
                            break;

                        #region Finding Noun Modifier first row in Attribute List File
                        ALFNMFoundRowNo = 0;
                        for (int afRowNo = 1; afRowNo <= ALFMaxRows; afRowNo++)
                        {
                            if (iNoun.ToUpper() == wsALF.Cells[afRowNo, 0].StringValue.Trim().ToUpper())
                            {
                                if (iModifier.ToUpper() == wsALF.Cells[afRowNo, 1].StringValue.Trim().ToUpper())
                                {
                                    ALFNMFoundRowNo = afRowNo;
                                    break;
                                }
                            }
                        }
                        #endregion

                        #region Adding Attribute Value Column Nos. from input file whose Attribute match to Attribute in Attribute List file To List
                        if (ALFNMFoundRowNo > 0)
                        {
                            IFAVColNoList.Clear();
                            for (int iANCol = 3; iANCol <= IFMaxColumns; iANCol += 2)
                            {
                                iAttributeName = wsIF.Cells[iRow, iANCol].StringValue.Trim();

                                if (string.IsNullOrEmpty(iAttributeName))
                                    break;

                                iAttributeNameFoundInAttributeList = false;
                                for (int afANCol = 2; afANCol <= ALFMaxColumns; afANCol++)
                                {
                                    if (string.IsNullOrEmpty(wsALF.Cells[ALFNMFoundRowNo, afANCol].StringValue.Trim()))
                                        break;

                                    if (iAttributeName.ToUpper() == wsALF.Cells[ALFNMFoundRowNo, afANCol].StringValue.Trim().ToUpper())
                                    {
                                        wsIF.Cells[iRow, iANCol].SetStyle(styleLightBlueFillColor);
                                        iAttributeNameFoundInAttributeList = true;
                                        break;
                                    }
                                }

                                if (iAttributeNameFoundInAttributeList)
                                    IFAVColNoList.Add(iANCol + 1);
                            }
                        }
                        #endregion

                        #region Iterating Input File Rows for Att. Value match 
                        IsDuplicateCounterUsed = false;
                        for (int irRow = iRow + 1; irRow <= IFMaxRows; irRow++)
                        {
                            if (string.IsNullOrEmpty(wsIF.Cells[irRow, ObservationColNo].StringValue.Trim()))
                            {
                                if (iNoun.ToUpper() == wsIF.Cells[irRow, 1].StringValue.Trim().ToUpper())
                                {
                                    if (iModifier.ToUpper() == wsIF.Cells[irRow, 2].StringValue.Trim().ToUpper())
                                    {
                                        AreAllColumnValuesMatching = false;
                                        for (int i = 0; i < IFAVColNoList.Count(); i++)
                                        {
                                            if ((wsIF.Cells[iRow, IFAVColNoList[i]].StringValue.Trim().ToUpper() ==
                                                 wsIF.Cells[irRow, IFAVColNoList[i]].StringValue.Trim().ToUpper()) &&
                                                !string.IsNullOrEmpty(wsIF.Cells[iRow, IFAVColNoList[i]].StringValue.Trim()) &&
                                                !string.IsNullOrEmpty(wsIF.Cells[irRow, IFAVColNoList[i]].StringValue.Trim()))
                                                AreAllColumnValuesMatching = true;
                                            else
                                            {
                                                AreAllColumnValuesMatching = false;
                                                break;
                                            }
                                        }

                                        if (AreAllColumnValuesMatching)
                                        {
                                            for (int i = 0; i < IFAVColNoList.Count(); i++)
                                                wsIF.Cells[irRow, IFAVColNoList[i] - 1].SetStyle(styleLightBlueFillColor);
                                            wsIF.Cells[iRow, ObservationColNo].PutValue("Exact Duplicate Set " + DuplicateSetCounter);
                                            wsIF.Cells[irRow, ObservationColNo].PutValue("Exact Duplicate Set " + DuplicateSetCounter);
                                            IsDuplicateCounterUsed = true;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        if (IsDuplicateCounterUsed)
                            DuplicateSetCounter++;
                    }
                }
                #endregion

                #region Save and Download the file
                wsIF.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);
                #endregion

                if (File.Exists(UploadedInputFilepath))
                    File.Delete(UploadedInputFilepath);

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
