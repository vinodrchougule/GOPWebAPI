using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Http;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/AttributeValueConsistencyCheck")]
    public class AttributeValueConsistencyCheckController : ApiController
    {
        #region Validate the input uploaded file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
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
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "SL NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SL No' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 1].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 2].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Noun'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 3].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'Modifier'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 4].StringValue.Trim().ToUpper() != "ATTRIBUTE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fifth Column with heading 'Attribute'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 5].StringValue.Trim().ToUpper() != "VALUE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Sixth Column with heading 'Value'  not found in input file first worksheet. Please select a valid file.");
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

        #region Validate the UOM uploaded file
        [HttpPost]
        [Route("ValidateUOMFile")]
        public HttpResponseMessage ValidateUOMFile(string FileName)
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
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid UOM file of xlsx format only");

                Workbook wbUOM = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUOM.Open(UploadedFilepath);
                Worksheet wsUOM = wbUOM.Worksheets[0];


                #region Check uploaded UOM file first worksheet has columns as per template
                if (wsUOM.Cells[0, 0].StringValue.Trim().ToUpper() != "UOM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'UOM' not found in UOM file first worksheet. Please select a valid file.");
                }

                if (wsUOM.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "UOM File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "UOM file validated successfully.");
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

        #region Fetch the unique UOMs from the UOM file
        [HttpGet]
        [Route("FetchUniqueUOMs")]
        public HttpResponseMessage FetchUniqueUOMs(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);
            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);
                Workbook wbUOM = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUOM.Open(UploadedFilepath);
                Worksheet wsUOM = wbUOM.Worksheets[0];
                DataTable dtUOM = wsUOM.Cells.ExportDataTable(1, 0, wsUOM.Cells.Rows.Count, 1);
                var uniqueUOMs = dtUOM.AsEnumerable().Select(r => r.Field<string>("Column1")).Where(u => !string.IsNullOrEmpty(u)).Distinct();
                return Request.CreateResponse(HttpStatusCode.OK, uniqueUOMs);
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Check the attribute value consistency and return the result file path
        [HttpPost]
        [Route("CheckAttributeValueConsistency")]
        public HttpResponseMessage CheckAttributeValueConsistency(AttributeValueConsistencyCheckModel model)
        {
            //Check Model State
            if (!ModelState.IsValid)
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

            string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + model.InputFileName);
            string UploadedUOMFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + model.UOMFileName);
            string AttributeValue = string.Empty, UOM = string.Empty;
            char aChar;

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                int LastColumnNo = wsIF.Cells.MaxColumn;
                int ObservationColumnNo = LastColumnNo + 1;
                int imaxRows = wsIF.Cells.MaxRow, imaxColumns = wsIF.Cells.MaxColumn;
                wsIF.Cells[0, ObservationColumnNo].PutValue("Observation(s)");

                Workbook wbUOM = new Workbook();
                Aspose.Cells.License l1 = new Aspose.Cells.License();
                l1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbUOM.Open(UploadedUOMFilepath);
                Worksheet wsUOM = wbUOM.Worksheets[0];
                int umaxRows = wsUOM.Cells.MaxRow;

                string[] UOMStringArray;
                bool IsObservationWritten = false;
                string AttributeValueWithOnlyAlphabets = string.Empty;
                string[] XSplittedArray;
                bool ZeroLengthStringFound = false;
                string MultipleValuesSeparator = model.MultipleValuesSeparator;
                string[] AttributeValueArray;

                for (int iRow = 1; iRow <= imaxRows; iRow++)
                {
                    AttributeValue = wsIF.Cells[iRow, 5].StringValue.Trim();

                    AttributeValueArray = AttributeValue.Split(new string[] { MultipleValuesSeparator }, StringSplitOptions.None);

                    foreach (string av in AttributeValueArray)
                    {
                        AttributeValue = av;

                        for (int uRow = 1; uRow <= umaxRows; uRow++)
                        {
                            UOM = wsUOM.Cells[uRow, 0].StringValue.Trim();

                            #region 1. Check spacing before UOM
                            if (AttributeValue.ToUpper().Contains(UOM.ToUpper()))
                            {
                                if (model.CheckSpaceBeforeUOM.Trim().ToUpper() == "W")
                                {
                                    if (AttributeValue.IndexOf(UOM) > 0)
                                    {
                                        if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf(UOM) - 1, 1), out aChar))
                                        {
                                            if (Char.IsDigit(aChar))
                                            {
                                                if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("No space between UOM and Value"))
                                                {
                                                    if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "No space between UOM and Value");
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (model.CheckSpaceBeforeUOM.Trim().ToUpper() == "O")
                                {
                                    if (AttributeValue.IndexOf(UOM) > 0)
                                    {
                                        if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf(UOM) - 1, 1), out aChar))
                                        {
                                            if (Char.IsWhiteSpace(aChar))
                                            {
                                                if (AttributeValue.IndexOf(UOM) + 1 < AttributeValue.Length)
                                                {
                                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf(UOM) + 1, 1), out aChar))
                                                    {
                                                        if (aChar.ToString().ToUpper() == "O")      //TO
                                                            continue;
                                                    }
                                                }

                                                if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space between UOM and Value not required"))
                                                {
                                                    if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space between UOM and Value not required");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion
                        }

                        #region 2. Multiple Dimension Separator
                        if (model.IsMultipleDimensionSeparatorXChecked)         //X
                        {
                            if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "W")
                            {
                                #region Check space on both sides of Multiple Dimension Separator X
                                if (AttributeValue.Trim().ToUpper().IndexOf('X') > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf('X') + 1) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('X') - 1, 1), out aChar))  //Left Character
                                    {
                                        if (!Char.IsWhiteSpace(aChar))
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator 'X'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator 'X'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('X') + 1, 1), out aChar))  //Right Character
                                            {
                                                if (!Char.IsWhiteSpace(aChar))
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator 'X'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator 'X'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "O")
                            {
                                #region Space should not be there on both sides of Multiple Value Separator X
                                if (AttributeValue.Trim().ToUpper().IndexOf('X') > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf('X') + 1) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('X') - 1, 1), out aChar))  //Left Character
                                    {
                                        if (Char.IsWhiteSpace(aChar))
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of multiple dimension separator 'X'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of multiple dimension separator 'X'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('X') + 1, 1), out aChar))  //Right Character
                                            {
                                                if (Char.IsWhiteSpace(aChar))
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of multiple dimension separator 'X'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of multiple dimension separator 'X'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                        }

                        if (model.IsMultipleDimensionSeparatorFWDSlashChecked)         //Forward Slash
                        {
                            if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "W")
                            {
                                #region Check space on both sides of Multiple Dimension Separator Forward Slash
                                if (AttributeValue.Trim().ToUpper().IndexOf('/') > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf('/') + 1) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('/') - 1, 1), out aChar))  //Left Character
                                    {
                                        if (!Char.IsWhiteSpace(aChar) && aChar != '+' && aChar != '-')
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator '/'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator '/'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('/') + 1, 1), out aChar))  //Right Character
                                            {
                                                if (!Char.IsWhiteSpace(aChar) && aChar != '+' && aChar != '-')
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator '/'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator '/'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "O")
                            {
                                #region Space should not be there on both sides of Multiple Value Separator Forward Slash
                                if (AttributeValue.Trim().ToUpper().IndexOf('/') > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf('/') + 1) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('/') - 1, 1), out aChar))  //Left Character
                                    {
                                        if (Char.IsWhiteSpace(aChar))
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of multiple dimension separator '/'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of multiple dimension separator '/'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('/') + 1, 1), out aChar))  //Right Character
                                            {
                                                if (Char.IsWhiteSpace(aChar))
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of multiple dimension separator '/'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of multiple dimension separator '/'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                        }

                        if (model.IsMultipleDimensionSeparatorTOChecked)         //TO
                        {
                            if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "W")
                            {
                                #region Check space on both sides of Multiple Dimension Separator TO
                                if (AttributeValue.Trim().ToUpper().IndexOf("TO") > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf("TO") + 2) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf("TO") - 1, 1), out aChar))  //Left Character
                                    {
                                        if (!Char.IsWhiteSpace(aChar))
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator 'TO'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator 'TO'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf("TO") + 2, 1), out aChar))  //Right Character
                                            {
                                                if (!Char.IsWhiteSpace(aChar))
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator 'TO'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator 'TO'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "O")
                            {
                                #region Space should not be there on both sides of Multiple Value Separator TO
                                if (AttributeValue.ToUpper().IndexOf("TO") > 0 &&
                                    ((AttributeValue.ToUpper().IndexOf("TO") + 2) != AttributeValue.Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.ToUpper().Substring(AttributeValue.ToUpper().IndexOf("TO") - 1, 1), out aChar))  //Left Character
                                    {
                                        if (Char.IsWhiteSpace(aChar))
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of multiple dimension separator 'TO'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of multiple dimension separator 'TO'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.ToUpper().Substring(AttributeValue.ToUpper().IndexOf("TO") + 2, 1), out aChar))  //Right Character
                                            {
                                                if (Char.IsWhiteSpace(aChar))
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of multiple dimension separator 'TO'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of multiple dimension separator 'TO'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                        }

                        if (model.IsMultipleDimensionSeparatorHyphenChecked)         //Hyphen
                        {
                            if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "W")
                            {
                                #region Check space on both sides of Multiple Dimension Separator Hyphen
                                if (AttributeValue.Trim().ToUpper().IndexOf('-') > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf('-') + 1) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('-') - 1, 1), out aChar))  //Left Character
                                    {
                                        if (!Char.IsWhiteSpace(aChar) && aChar != '/')
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator '-'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator '-'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('-') + 1, 1), out aChar))  //Right Character
                                            {
                                                if (!Char.IsWhiteSpace(aChar) && aChar != '/' && AttributeValue.Substring(AttributeValue.IndexOf('-') - 1, 1) != "/")
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of multiple dimension separator '-'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of multiple dimension separator '-'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (model.CheckSpaceForMultipleDimensionSeparator.Trim().ToUpper() == "O")
                            {
                                #region Space should not be there on both sides of Multiple Dimension Separator Hyphen
                                if (AttributeValue.Contains(" - ") || AttributeValue.Contains(" -") || AttributeValue.Contains("- "))
                                {
                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of multiple dimension separator '-'"))
                                    {
                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of multiple dimension separator '-'");
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        #region 3. Range Values Separator
                        if (model.IsRangeValuesSeparatorTOChecked)               //TO
                        {
                            if (model.CheckSpaceForRangeValuesSeparator.Trim().ToUpper() == "W")
                            {
                                #region Check space on both sides of Range Values Separator TO
                                if (AttributeValue.Trim().ToUpper().IndexOf("TO") > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf("TO") + 2) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf("TO") - 1, 1), out aChar))  //Left Character
                                    {
                                        if (!Char.IsWhiteSpace(aChar))
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of range values separator 'TO'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of range values separator 'TO'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf("TO") + 2, 1), out aChar))  //Right Character
                                            {
                                                if (!Char.IsWhiteSpace(aChar))
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of range values separator 'TO'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of range values separator 'TO'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (model.CheckSpaceForRangeValuesSeparator.Trim().ToUpper() == "O")
                            {
                                #region Space should not be there on both sides of Range Values Separator TO
                                if (AttributeValue.ToUpper().IndexOf("TO") > 0 &&
                                    ((AttributeValue.ToUpper().IndexOf("TO") + 2) != AttributeValue.Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.ToUpper().Substring(AttributeValue.ToUpper().IndexOf("TO") - 1, 1), out aChar))  //Left Character
                                    {
                                        if (Char.IsWhiteSpace(aChar))
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of range values separator 'TO'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of range values separator 'TO'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.ToUpper().Substring(AttributeValue.ToUpper().IndexOf("TO") + 2, 1), out aChar))  //Right Character
                                            {
                                                if (Char.IsWhiteSpace(aChar))
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of range values separator 'TO'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of range values separator 'TO'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                        }

                        if (model.IsRangeValuesSeparatorHyphenChecked)           //Hyphen
                        {
                            if (model.CheckSpaceForRangeValuesSeparator.Trim().ToUpper() == "W")
                            {
                                #region Check space on both sides of Range Values Separator Hyphen
                                if (AttributeValue.Trim().ToUpper().IndexOf('-') > 0 &&
                                    ((AttributeValue.Trim().ToUpper().IndexOf('-') + 1) != AttributeValue.Trim().Length))     //Not at the beginning and at the end
                                {
                                    if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('-') - 1, 1), out aChar))  //Left Character
                                    {
                                        if (!Char.IsWhiteSpace(aChar) && aChar != '/')
                                        {
                                            if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of range values separator '-'"))
                                            {
                                                if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of range values separator '-'");
                                            }
                                        }
                                        else
                                        {
                                            if (Char.TryParse(AttributeValue.Substring(AttributeValue.IndexOf('-') + 1, 1), out aChar))  //Right Character
                                            {
                                                if (!Char.IsWhiteSpace(aChar) && aChar != '/' && AttributeValue.Substring(AttributeValue.IndexOf('-') - 1, 1) != "/")
                                                {
                                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space has to be there on both sides of range values separator '-'"))
                                                    {
                                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space has to be there on both sides of range values separator '-'");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (model.CheckSpaceForRangeValuesSeparator.Trim().ToUpper() == "O")
                            {
                                #region Space should not be there on both sides of Range Values Separator Hyphen
                                if (AttributeValue.Contains(" - ") || AttributeValue.Contains(" -") || AttributeValue.Contains("- "))
                                {
                                    if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Space not required on both sides of range values separator '-'"))
                                    {
                                        if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Space not required on both sides of range values separator '-'");
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        #region Check same UOMs on both sides of X
                        if (AttributeValue.ToUpper().Contains('X'))
                        {
                            AttributeValueWithOnlyAlphabets = Regex.Replace(AttributeValue, @"[^a-zA-Z]+", string.Empty);
                            XSplittedArray = AttributeValueWithOnlyAlphabets.ToUpper().Split('X');
                            IsObservationWritten = false;
                            for (int i = 0; i < XSplittedArray.Count(); i++)
                            {
                                if (XSplittedArray[i].Trim().Length == 0)
                                    break;

                                for (int j = i + 1; j < XSplittedArray.Count(); j++)
                                {
                                    if (XSplittedArray[j].Trim().Length == 0)
                                    {
                                        ZeroLengthStringFound = true;
                                        break;
                                    }

                                    if (XSplittedArray[i].Trim().ToUpper() != XSplittedArray[j].Trim().ToUpper())
                                    {
                                        if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Same UOMs should be there on both sides of 'X'"))
                                        {
                                            if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Same UOMs should be there on both sides of 'X'");
                                            IsObservationWritten = true;
                                            break;
                                        }
                                    }
                                }

                                if (ZeroLengthStringFound || IsObservationWritten)
                                    break;
                            }
                        }
                        #endregion

                        #region 4. Check the Order of UOMs in Attribute Value with order in list
                        if (model.SelectedOrderedUOMList != null && model.SelectedOrderedUOMList.Count() > 0)
                        {
                            UOMStringArray = model.SelectedOrderedUOMList.ToArray();
                            string[] UOMArray = model.SelectedOrderedUOMList.ToArray();
                            IsObservationWritten = false;
                            for (int i = 0; i < UOMStringArray.Count(); i++)
                            {
                                if (AttributeValue.ToUpper().IndexOf(UOMArray[i].ToUpper()) > -1)
                                {
                                    for (int j = i + 1; j < UOMStringArray.Count(); j++)
                                    {
                                        if (AttributeValue.ToUpper().IndexOf(UOMArray[j].ToUpper()) > 0)
                                        {
                                            if (AttributeValue.ToUpper().IndexOf(UOMArray[i].ToUpper()) > AttributeValue.ToUpper().IndexOf(UOMArray[j].ToUpper()))
                                            {
                                                if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Order of UOMs needs to be in specified order"))
                                                {
                                                    if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                        wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                                    wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Order of UOMs needs to be in specified order");
                                                }
                                                IsObservationWritten = true;
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (IsObservationWritten)
                                    break;
                            }
                        }
                        #endregion

                        #region 5. Check conversion of Volts/Amps to KiloVolts/KiloAmps required
                        if (AttributeValue.Trim().Length > 1 &&
                            (AttributeValue.ToUpper().Contains('V') || AttributeValue.ToUpper().Contains("VAC") ||
                             AttributeValue.ToUpper().Contains("VDC") || AttributeValue.ToUpper().Contains('A') ||
                             AttributeValue.ToUpper().Contains("VA")))
                        {
                            if (model.IsConversionOfVAChecked)
                            {
                                decimal number;
                                var units = new[] { "V", "VAC", "VDC", "A", "VA" };
                                var splitnumber = AttributeValue.ToUpper().Split(units, StringSplitOptions.RemoveEmptyEntries);
                                if (decimal.TryParse(splitnumber[0], out number))
                                {
                                    if (number >= Convert.ToDecimal(1000))
                                    {
                                        if (!wsIF.Cells[iRow, 6].StringValue.Trim().Contains("Conversion of Volts/Amps to KiloVolts/KiloAmps required."))
                                        {
                                            if (wsIF.Cells[iRow, 6].StringValue.Trim().Length > 0)
                                                wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + ",");
                                            wsIF.Cells[iRow, 6].PutValue(wsIF.Cells[iRow, 6].StringValue.Trim() + "Conversion of Volts/Amps to KiloVolts/KiloAmps required.");
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                }

                #region Save and Download the file
                wsIF.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + model.UploadedInputFileName;
                wbIF.Save(filename);
                #endregion

                if (File.Exists(UploadedInputFilepath))
                    File.Delete(UploadedInputFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedInputFilepath))
                    File.Delete(UploadedInputFilepath);

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
