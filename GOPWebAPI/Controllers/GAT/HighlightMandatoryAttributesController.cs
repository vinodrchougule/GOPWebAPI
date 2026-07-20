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
using GOPWebAPI.Models.GAT_Models;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/HighlightMandatoryAttributes")]
    public class HighlightMandatoryAttributesController : ApiController
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
                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "SNO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SNo' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 1].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Noun' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 2].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                string hdr;
                int attNo = 1;
                for (int aCtr = 1; aCtr <= 100; aCtr += 2)
                {
                    hdr = wsIF.Cells[0, aCtr + 2].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 3).ToString() + "] must contain 'Attribute " + attNo.ToString() + "' .Please download format from Template.");
                    }

                    hdr = wsIF.Cells[0, aCtr + 3].StringValue.Trim().ToUpper();
                    if (hdr != "ATTRIBUTE VALUE " + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (aCtr + 4).ToString() + "] must contain 'Attribute Value " + attNo.ToString() + "' .Please download format from Template.");
                    }

                    attNo++;
                }

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

        #region Validate the dictionary uploaded file
        [HttpPost]
        [Route("ValidateDictionaryFile")]
        public HttpResponseMessage ValidateDictionaryFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid dictionary file of xlsx format only");

                Workbook wbDF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbDF.Open(UploadedFilepath);
                Worksheet wsDF = wbDF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsDF.Cells[0, 0].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Noun' not found in dictionary file first worksheet. Please select a valid file.");
                }

                if (wsDF.Cells[0, 1].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Modifier' not found in dictionary file first worksheet. Please select a valid file.");
                }

                if (wsDF.Cells[0, 2].StringValue.Trim().ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Attribute' not found in dictionary file first worksheet. Please select a valid file.");
                }

                if (wsDF.Cells[0, 3].StringValue.Trim().ToUpper() != "MANDATORY OR OPTIONAL (M OR O)")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'Mandatory or Optional (M or O)' not found in dictionary file first worksheet. Please select a valid file.");
                }

                if (wsDF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Dictionary File first Worksheet has no data rows.");
                }

                for (int row = 1; row <= wsDF.Cells.MaxRow; row++)
                {
                    if (wsDF.Cells[row,3].StringValue.Trim().ToLower()!="m" && 
                        wsDF.Cells[row, 3].StringValue.Trim().ToLower() != "o")
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Dictionary File 'Mandatory or Optional (M or O)' column should have only M or O character. Row No.:" + row.ToString());
                    }
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Dictionary file validated successfully.");
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

        #region Highlight Mandatory Attributes
        [HttpPost]
        [Route("HighlightMandatoryAttributes")]
        public HttpResponseMessage HighlightMandatoryAttributes(string InputFileName,string DictionaryFileName)
        {
            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
            string DictionaryFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + DictionaryFileName);

            try
            {
                #region Add Mandatory Attributes from Dictionary to List
                List<Models.GAT_Models.NounModifierAttribute> mandatoryAttributesList = new List<Models.GAT_Models.NounModifierAttribute>();

                Workbook wbDF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbDF.Open(DictionaryFilepath);
                Worksheet wsDF = wbDF.Worksheets[0];
                
                for(int row=1; row<=wsDF.Cells.MaxRow; row++)
                {
                    if (wsDF.Cells[row,3].StringValue.Trim().ToLower() == "m")
                    {
                        Models.GAT_Models.NounModifierAttribute nounModifierAttribute = new Models.GAT_Models.NounModifierAttribute();
                        nounModifierAttribute.Noun = wsDF.Cells[row, 0].StringValue.Trim();
                        nounModifierAttribute.Modifier = wsDF.Cells[row, 1].StringValue.Trim();
                        nounModifierAttribute.Attribute = wsDF.Cells[row, 2].StringValue.Trim();
                        mandatoryAttributesList.Add(nounModifierAttribute);
                    }
                }
                #endregion

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l1 = new Aspose.Cells.License();
                l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(InputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Create style for highlighting
                Aspose.Cells.Style styleLightRedFillColor = wsIF.Cells[1, 3].GetStyle();

                //Setting fill-color style
                styleLightRedFillColor.ForegroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                styleLightRedFillColor.BackgroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                styleLightRedFillColor.Pattern = BackgroundType.VerticalStripe;
                styleLightRedFillColor.Font.Color = System.Drawing.Color.Black;
                #endregion

                #region Navigate all the rows of input file and highligt the attribute if it is mandatory
                for (int row=1; row<= wsIF.Cells.MaxRow; row++)
                {
                    for (int col = 3; col<=wsIF.Cells.MaxColumn; col+=2)
                    {
                        if (string.IsNullOrEmpty(wsIF.Cells[row, col].StringValue))
                            break;

                        if (mandatoryAttributesList.Any(x => 
                                                            x.Noun.ToLower() == wsIF.Cells[row,1].StringValue.ToLower().Trim() &&
                                                            x.Modifier.ToLower() == wsIF.Cells[row, 2].StringValue.ToLower().Trim() &&
                                                            x.Attribute.ToLower() == wsIF.Cells[row,col].StringValue.ToLower().Trim()))
                        {
                            wsIF.Cells[row,col].SetStyle(styleLightRedFillColor);
                        }
                    }
                }
                #endregion

                wsIF.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + InputFileName;
                wbIF.Save(filename);

                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

                if (File.Exists(DictionaryFilepath))
                    File.Delete(DictionaryFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
            }
            catch (Exception ex)
            {
                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

                if (File.Exists(DictionaryFilepath))
                    File.Delete(DictionaryFilepath);

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
