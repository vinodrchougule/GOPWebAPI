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
    [RoutePrefix("api/ProductionVsQACheck")]
    public class ProductionVsQACheckController : ApiController
    {
        #region Validate the input uploaded file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            if (!File.Exists(UploadedFilepath))
                return Request.CreateResponse(HttpStatusCode.BadRequest, "File not found");
            else
            {
                string extension = Path.GetExtension(UploadedFilepath);
                if (extension != ".xlsx")
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Invalid file format. Please upload an Excel file.");
                else
                {
                    try
                    {
                        Workbook workbook = new Workbook();
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        workbook.Open(UploadedFilepath);

                        if(workbook.Worksheets.Count < 2)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File should have at least 2 worksheets.");
                        }

                        Worksheet wsProduction = workbook.Worksheets[0];
                        Worksheet wsDelivered = workbook.Worksheets[1];

                        #region Validating Production worksheet
                        if (wsProduction.Cells.Rows.Count <= 1)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Production Worksheet has no data rows.");
                        }

                        int pcAttributeName1ColNo = 0, pcAttributeValue1ColNo = 0;
                        bool IDColExists = false, ReferenceURL1ColExists = false, NounColExists = false, ModifierColExists = false, UNSPSCCodeColExists = false, UNSPSCDescriptionColExists = false;
                        string UnMatchedColHeader = "";

                        for (int ColCounter = 0; ColCounter <= wsProduction.Cells.MaxColumn; ColCounter++)
                        {
                            if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() != wsDelivered.Cells[0, ColCounter].StringValue.ToUpper())
                            {
                                UnMatchedColHeader = wsProduction.Cells[0, ColCounter].StringValue;
                                break;
                            }

                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "ID")
                                IDColExists = true;

                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "REFERENCE URL 1")
                                ReferenceURL1ColExists = true;

                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "NOUN")
                                NounColExists = true;

                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "MODIFIER")
                                ModifierColExists = true;

                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "UNSPSC CODE")
                                UNSPSCCodeColExists = true;


                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "UNSPSC DESCRIPTION")
                                UNSPSCDescriptionColExists = true;

                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "ATTRIBUTE NAME 1")
                                pcAttributeName1ColNo = ColCounter;

                            if (wsProduction.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "ATTRIBUTE VALUE 1")
                                pcAttributeValue1ColNo = ColCounter;
                        }

                        if (UnMatchedColHeader.Trim().Length > 0)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "The column heading '" + UnMatchedColHeader + "' is not matching from 'Production' and 'Delivered' worksheets.Please download template and fill-up the data.");
                        }

                        if (!IDColExists)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'ID' column not found in 'Production' worksheet.Please download template and fill-up the data.");
                        }

                        if (!ReferenceURL1ColExists)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Reference URL 1' column not found in 'Production' worksheet.Please download template and fill-up the data.");
                        }

                        if (!UNSPSCCodeColExists)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Code' column not found in 'Production' worksheet.Please download template and fill-up the data.");
                        }

                        if (!UNSPSCDescriptionColExists)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'UNSPSC Description' column not found in 'Production' worksheet.Please download template and fill-up the data.");
                        }

                        if (!NounColExists)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Noun' column not found in 'Production' worksheet.Please download template and fill-up the data.");
                        }

                        if (!ModifierColExists)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "'Modifier' column not found in 'Production' worksheet.Please download template and fill-up the data.");
                        }

                        #region Checking Attribute Names and Values from 'Production' worksheet
                        string hdr;
                        int attNo = 1;
                        for (int aCtr = 1; aCtr <= 100; aCtr += 2)
                        {
                            hdr = wsProduction.Cells[0, pcAttributeName1ColNo].StringValue.Trim().ToUpper();
                            
                            if (hdr != "ATTRIBUTE NAME " + attNo.ToString())
                            {
                                File.Delete(UploadedFilepath);
                                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (pcAttributeName1ColNo).ToString() + "] must contain 'Attribute Name " + attNo.ToString() + "' in Production worksheet first row.Please download format from Template.");
                            }

                            hdr = wsProduction.Cells[0, pcAttributeValue1ColNo].StringValue.Trim().ToUpper();
                            if (hdr != "ATTRIBUTE VALUE " + attNo.ToString())
                            {
                                File.Delete(UploadedFilepath);
                                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (pcAttributeValue1ColNo).ToString() + "] must contain 'Attribute Value " + attNo.ToString() + "' in Production worksheet first row.Please download format from Template.");
                            }

                            pcAttributeName1ColNo = pcAttributeName1ColNo + 2;
                            pcAttributeValue1ColNo = pcAttributeValue1ColNo + 2;
                            attNo++;
                        }
                        #endregion

                        #endregion

                        #region Validating Delivered worksheet
                        if (wsDelivered.Cells.Rows.Count <= 1)
                        {
                            File.Delete(UploadedFilepath);
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Delivered Worksheet has no data rows.");
                        }

                        int dAttributeName1ColNo = 0, dAttributeValue1ColNo = 0;
                        IDColExists = false; ReferenceURL1ColExists = false; NounColExists = false; ModifierColExists = false; UNSPSCCodeColExists = false; UNSPSCDescriptionColExists = false;
                        for (int ColCounter = 0; ColCounter <= wsDelivered.Cells.MaxColumn; ColCounter++)
                        {
                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "ID")
                                IDColExists = true;

                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "REFERENCE URL 1")
                                ReferenceURL1ColExists = true;

                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "NOUN")
                                NounColExists = true;

                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "MODIFIER")
                                ModifierColExists = true;

                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "UNSPSC CODE")
                                UNSPSCCodeColExists = true;

                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "UNSPSC DESCRIPTION")
                                UNSPSCDescriptionColExists = true;

                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "ATTRIBUTE NAME 1")
                                dAttributeName1ColNo = ColCounter;

                            if (wsDelivered.Cells[0, ColCounter].StringValue.Trim().ToUpper() == "ATTRIBUTE VALUE 1")
                                dAttributeValue1ColNo = ColCounter;
                        }

                        attNo = 1;
                        for (int aCtr = 1; aCtr <= 100; aCtr += 2)
                        {
                            hdr = wsDelivered.Cells[0, dAttributeName1ColNo].StringValue.Trim().ToUpper();
                            if (hdr != "ATTRIBUTE NAME " + attNo.ToString())
                            {
                                File.Delete(UploadedFilepath);
                                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (dAttributeName1ColNo).ToString() + "] must contain 'Attribute Name " + attNo.ToString() + "' in Delivered worksheet first row. Please download format from Template.");
                            }
                            hdr = wsDelivered.Cells[0, dAttributeValue1ColNo].StringValue.Trim().ToUpper();
                            if (hdr != "ATTRIBUTE VALUE " + attNo.ToString())
                            {
                                File.Delete(UploadedFilepath);
                                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell[1," + (dAttributeValue1ColNo).ToString() + "] must contain 'Attribute Value " + attNo.ToString() + "' in Delivered worksheet first row.Please download format from Template.");
                            }
                            dAttributeName1ColNo = dAttributeName1ColNo + 2;
                            dAttributeValue1ColNo = dAttributeValue1ColNo + 2;
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
            }
        }
        #endregion

        #region Check For Differences and Write the differences to Output file
        [HttpPost]
        [Route("CheckForDifferencesAndWriteToOutputFile")]
        public HttpResponseMessage CheckForDifferencesAndWriteToOutputFile(string InputFileName, string UploadedInputFileName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedInputFilepath);

                Worksheet wsProduction = wbIF.Worksheets[0];
                Worksheet wsDelivered = wbIF.Worksheets[1];

                bool IsCheckDifferenceReportWorksheetExist = false, IsSummaryWorksheetExist = false;
                foreach (Worksheet s in wbIF.Worksheets)
                {
                    if (s.Name.Trim().ToUpper() == "ERROR REPORT")
                    {
                        IsCheckDifferenceReportWorksheetExist = true;
                        break;
                    }

                    if (s.Name.Trim().ToUpper() == "SUMMARY")
                    {
                        IsSummaryWorksheetExist = true;
                        break;
                    }
                }

                if (!IsCheckDifferenceReportWorksheetExist)
                    wbIF.Worksheets.Add("Error Report");

                Worksheet wsErrorReport = wbIF.Worksheets[2];

                if (!IsSummaryWorksheetExist)
                    wbIF.Worksheets.Add("Summary");

                Worksheet wsSummaryReport = wbIF.Worksheets[3];

                Aspose.Cells.Style styleLightRedFillColor = wsErrorReport.Cells[1, 0].GetStyle();

                //Copying 'Delivered' worksheet contents to 'Error Report' 
                wsErrorReport.Copy(wsDelivered);

                #region Finding out the required column nos. from 'Production' worksheet
                int pcIDColNo = 0, pcReferenceURL1ColNo = 0, pcNounColNo = 0, pcModifierColNo = 0, pcAttributeValue1ColNo = 0, pcUNSPSCCodeColNo = 0, pcUNSPSCDescriptionColNo = 0;
                for (int ColCounter = 0; ColCounter <= wsProduction.Cells.MaxColumn; ColCounter++)
                {
                    if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() == "ID")
                        pcIDColNo = ColCounter;

                    if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() == "REFERENCE URL 1")
                        pcReferenceURL1ColNo = ColCounter;

                    if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() == "NOUN")
                        pcNounColNo = ColCounter;

                    if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() == "MODIFIER")
                        pcModifierColNo = ColCounter;

                    if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() == "ATTRIBUTE VALUE 1")
                        pcAttributeValue1ColNo = ColCounter;

                    if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() == "UNSPSC CODE")
                        pcUNSPSCCodeColNo = ColCounter;

                    if (wsProduction.Cells[0, ColCounter].StringValue.ToUpper() == "UNSPSC DESCRIPTION")
                        pcUNSPSCDescriptionColNo = ColCounter;

                    if (pcIDColNo > 0 && pcReferenceURL1ColNo > 0 && pcNounColNo > 0 && pcModifierColNo > 0 && pcAttributeValue1ColNo > 0 && pcUNSPSCCodeColNo > 0 && pcUNSPSCDescriptionColNo > 0)
                        break;
                }
                #endregion

                #region Finding out the required column nos. from 'Error Report' worksheet
                int cdrIDColNo = 0, cdrReferenceURL1ColNo = 0, cdrNounColNo = 0, cdrModifierColNo = 0, cdrAttributeValue1ColNo = 0, cdrUNSPSCCodeColNo = 0, cdrUNSPSCDescriptionColNo = 0;
                for (int ColCounter = 0; ColCounter <= wsDelivered.Cells.MaxColumn; ColCounter++)
                {
                    if (wsErrorReport.Cells[0, ColCounter].StringValue.ToUpper() == "ID")
                        cdrIDColNo = ColCounter;

                    if (wsErrorReport.Cells[0, ColCounter].StringValue.ToUpper() == "REFERENCE URL 1")
                        cdrReferenceURL1ColNo = ColCounter;

                    if (wsErrorReport.Cells[0, ColCounter].StringValue.ToUpper() == "NOUN")
                        cdrNounColNo = ColCounter;

                    if (wsErrorReport.Cells[0, ColCounter].StringValue.ToUpper() == "MODIFIER")
                        cdrModifierColNo = ColCounter;

                    if (wsErrorReport.Cells[0, ColCounter].StringValue.ToUpper() == "ATTRIBUTE VALUE 1")
                        cdrAttributeValue1ColNo = ColCounter;

                    if (wsErrorReport.Cells[0, ColCounter].StringValue.ToUpper() == "UNSPSC CODE")
                        cdrUNSPSCCodeColNo = ColCounter;

                    if (wsErrorReport.Cells[0, ColCounter].StringValue.ToUpper() == "UNSPSC DESCRIPTION")
                        cdrUNSPSCDescriptionColNo = ColCounter;

                    if (cdrIDColNo > 0 && cdrReferenceURL1ColNo > 0 && cdrNounColNo > 0 && cdrModifierColNo > 0 && cdrAttributeValue1ColNo > 0 && cdrUNSPSCCodeColNo > 0 && cdrUNSPSCDescriptionColNo > 0)
                        break;
                }
                #endregion

                #region Checking the data rows from 'Error Report' worksheet to 'Production' and filling unmatched cells with light red colour
                //Setting fill-color style
                styleLightRedFillColor.ForegroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                styleLightRedFillColor.BackgroundColor = System.Drawing.Color.FromArgb(0, 250, 191, 143);
                styleLightRedFillColor.Pattern = BackgroundType.VerticalStripe;
                styleLightRedFillColor.Font.Color = System.Drawing.Color.Black;

                //Highlighting the cells not matching the contents
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    for (int cdrColNo = 0; cdrColNo <= wsErrorReport.Cells.MaxColumn; cdrColNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrColNo].StringValue.Trim().ToUpper() != wsProduction.Cells[cdrRowNo, cdrColNo].StringValue.Trim().ToUpper())
                        {
                            if (cdrColNo == cdrUNSPSCCodeColNo)
                            {
                                if (wsErrorReport.Cells[cdrRowNo, cdrUNSPSCCodeColNo].StringValue.Trim().Length == 8 &&
                                    wsProduction.Cells[cdrRowNo, pcUNSPSCCodeColNo].StringValue.Trim().Length == 8
                                    )
                                {
                                    if (wsErrorReport.Cells[cdrRowNo, cdrUNSPSCCodeColNo].StringValue.Trim().Substring(0, 6) ==
                                        wsProduction.Cells[cdrRowNo, pcUNSPSCCodeColNo].StringValue.Trim().Substring(0, 6))           //do not highlight the UNSPSC Code cells if matches till class level 
                                    {
                                        cdrColNo++;
                                        continue;
                                    }
                                }
                            }

                            wsErrorReport.Cells[cdrRowNo, cdrColNo].SetStyle(styleLightRedFillColor);
                        }
                    }
                }
                #endregion

                #region Checking 'Reference URL 1' column to 'Production Completed' worksheet
                int cdrReferenceURL1MatchRemarksColNo = 0;
                bool IsRefURL1Exist = false;
                for (int pcRowNo = 1; pcRowNo <= wsProduction.Cells.MaxRow; pcRowNo++)
                {
                    if (wsProduction.Cells[pcRowNo, pcReferenceURL1ColNo].StringValue.Trim().Length > 0)
                    {
                        IsRefURL1Exist = true;
                        break;
                    }
                }

                if (!IsRefURL1Exist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrReferenceURL1ColNo].StringValue.Trim().Length > 0)
                        {
                            IsRefURL1Exist = true;
                            break;
                        }
                    }
                }

                cdrReferenceURL1MatchRemarksColNo = wsErrorReport.Cells.MaxColumn + 1;
                wsErrorReport.Cells[0, cdrReferenceURL1MatchRemarksColNo].PutValue("Reference URL 1 Match Remarks");
                if (IsRefURL1Exist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrReferenceURL1ColNo].StringValue.Trim().Length == 0 &&
                           wsProduction.Cells[cdrRowNo, pcReferenceURL1ColNo].StringValue.Trim().Length > 0)
                        {
                            wsErrorReport.Cells[cdrRowNo, cdrReferenceURL1MatchRemarksColNo].PutValue("Production URL Removed");
                        }
                    }
                }
                #endregion

                #region Checking 'UNSPSC' column to 'Production' worksheet
                bool IsUNSPSCExist = false;
                for (int pcRowNo = 1; pcRowNo <= wsProduction.Cells.MaxRow; pcRowNo++)
                {
                    if (wsProduction.Cells[pcRowNo, pcUNSPSCCodeColNo].StringValue.Trim().Length > 0)
                    {
                        IsUNSPSCExist = true;
                        break;
                    }
                }

                if (!IsUNSPSCExist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrUNSPSCCodeColNo].StringValue.Trim().Length > 0)
                        {
                            IsUNSPSCExist = true;
                            break;
                        }
                    }
                }

                int cdrIsUNSPSCMatchedColNo = wsErrorReport.Cells.MaxColumn + 1;
                wsErrorReport.Cells[0, cdrIsUNSPSCMatchedColNo].PutValue("Is UNSPSC Matched?");
                if (IsUNSPSCExist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsProduction.Cells[cdrRowNo, pcUNSPSCCodeColNo].StringValue.Trim().Length > 0)        //added to show as an error only for Production UNSPSC existing rows
                        {
                            if ((wsErrorReport.Cells[cdrRowNo, cdrUNSPSCCodeColNo].StringValue.Trim().ToUpper() !=
                                 wsProduction.Cells[cdrRowNo, pcUNSPSCCodeColNo].StringValue.Trim().ToUpper()) ||
                                (wsErrorReport.Cells[cdrRowNo, cdrUNSPSCDescriptionColNo].StringValue.Trim().ToUpper() !=
                                 wsProduction.Cells[cdrRowNo, pcUNSPSCDescriptionColNo].StringValue.Trim().ToUpper())
                               )
                            {
                                if (wsErrorReport.Cells[cdrRowNo, cdrUNSPSCCodeColNo].StringValue.Trim().Length == 8 &&
                                    wsProduction.Cells[cdrRowNo, pcUNSPSCCodeColNo].StringValue.Trim().Length == 8)
                                {
                                    if (wsErrorReport.Cells[cdrRowNo, cdrUNSPSCCodeColNo].StringValue.Trim().Substring(0, 6) ==
                                        wsProduction.Cells[cdrRowNo, pcUNSPSCCodeColNo].StringValue.Trim().Substring(0, 6))        //added to skip the error, if UNSPSC Code matches till class level -- (as DIW said)
                                        continue;
                                    else
                                        wsErrorReport.Cells[cdrRowNo, cdrIsUNSPSCMatchedColNo].PutValue("No");
                                }
                                else
                                    wsErrorReport.Cells[cdrRowNo, cdrIsUNSPSCMatchedColNo].PutValue("No");
                            }
                        }
                    }
                }
                #endregion

                #region Checking 'Noun' column to 'Production' worksheet
                bool IsNounExist = false;
                int cdrIsNounMatchedColNo = wsErrorReport.Cells.MaxColumn + 1;

                for (int pcRowNo = 1; pcRowNo <= wsProduction.Cells.MaxRow; pcRowNo++)
                {
                    if (wsProduction.Cells[pcRowNo, pcNounColNo].StringValue.Trim().Length > 0)
                    {
                        IsNounExist = true;
                        break;
                    }
                }

                if (!IsNounExist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrNounColNo].StringValue.Trim().Length > 0)
                        {
                            IsNounExist = true;
                            break;
                        }
                    }
                }

                wsErrorReport.Cells[0, cdrIsNounMatchedColNo].PutValue("Is Noun Matched?");
                if (IsNounExist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if ((wsErrorReport.Cells[cdrRowNo, cdrNounColNo].StringValue.Trim().ToUpper() !=
                             wsProduction.Cells[cdrRowNo, pcNounColNo].StringValue.Trim().ToUpper()))
                            wsErrorReport.Cells[cdrRowNo, cdrIsNounMatchedColNo].PutValue("No");
                    }
                }
                #endregion

                #region Checking 'Modifier' column to 'Production' worksheet
                int cdrIsModifierMatchedColNo = wsErrorReport.Cells.MaxColumn + 1;
                bool IsModifierExist = false;

                for (int pcRowNo = 1; pcRowNo <= wsProduction.Cells.MaxRow; pcRowNo++)
                {
                    if (wsProduction.Cells[pcRowNo, pcModifierColNo].StringValue.Trim().Length > 0)
                    {
                        IsModifierExist = true;
                        break;
                    }
                }

                if (!IsModifierExist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrModifierColNo].StringValue.Trim().Length > 0)
                        {
                            IsModifierExist = true;
                            break;
                        }
                    }
                }

                wsErrorReport.Cells[0, cdrIsModifierMatchedColNo].PutValue("Is Modifier Matched?");
                if (IsModifierExist)
                {
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrIsNounMatchedColNo].StringValue.Trim().Length == 0)
                        {
                            if ((wsErrorReport.Cells[cdrRowNo, cdrModifierColNo].StringValue.ToUpper() !=
                                 wsProduction.Cells[cdrRowNo, pcModifierColNo].StringValue.ToUpper()))
                                wsErrorReport.Cells[cdrRowNo, cdrIsModifierMatchedColNo].PutValue("No");
                        }
                    }
                }
                #endregion

                #region Checking 'Attribute Values' columns to 'Production' worksheet
                int AttributeValuesNotMatched = 0, cdrAttributeValueColNo = 0, pcAttributeValueColNo = 0;
                int cdrHowManyAttributeValuesNotMatchedColNo = wsErrorReport.Cells.MaxColumn + 1;
                wsErrorReport.Cells[0, cdrHowManyAttributeValuesNotMatchedColNo].PutValue("How Many Attribute Values Not Matched?");
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    if (wsErrorReport.Cells[cdrRowNo, cdrIsNounMatchedColNo].StringValue.Trim().Length == 0 &&
                        wsErrorReport.Cells[cdrRowNo, cdrIsModifierMatchedColNo].StringValue.Trim().Length == 0)
                    {
                        AttributeValuesNotMatched = 0;
                        cdrAttributeValueColNo = cdrAttributeValue1ColNo;
                        pcAttributeValueColNo = pcAttributeValue1ColNo;
                        for (int AttributeValueCounter = 1; AttributeValueCounter <= 50; AttributeValueCounter++)
                        {
                            if (wsErrorReport.Cells[cdrRowNo, cdrAttributeValueColNo].StringValue.Trim().Replace(" ", "").ToUpper() !=
                                wsProduction.Cells[cdrRowNo, pcAttributeValueColNo].StringValue.Trim().Replace(" ", "").ToUpper())
                                AttributeValuesNotMatched++;
                            cdrAttributeValueColNo = cdrAttributeValueColNo + 2;
                            pcAttributeValueColNo = pcAttributeValueColNo + 2;
                        }

                        if (AttributeValuesNotMatched > 0)
                            wsErrorReport.Cells[cdrRowNo, cdrHowManyAttributeValuesNotMatchedColNo].PutValue(AttributeValuesNotMatched);
                    }
                }
                #endregion

                #region Checking No. of Attributes Names, No. of Attribute Values Filled
                int NoOfAttributeNames = 0, cdrAttributeNameColNo = 0, NoOfAttributeValuesFilled = 0;
                int cdrNoOfAttributeNamesColNo = wsErrorReport.Cells.MaxColumn + 1;
                int cdrNoOfAttributeValuesFilledColNo = wsErrorReport.Cells.MaxColumn + 2;
                wsErrorReport.Cells[0, cdrNoOfAttributeNamesColNo].PutValue("No. of Attribute Names");
                wsErrorReport.Cells[0, cdrNoOfAttributeValuesFilledColNo].PutValue("No. of Attribute Values Filled");
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    NoOfAttributeNames = 0;
                    NoOfAttributeValuesFilled = 0;
                    cdrAttributeValueColNo = cdrAttributeValue1ColNo;
                    cdrAttributeNameColNo = cdrAttributeValue1ColNo - 1;

                    for (int AttributeNameCounter = 1; AttributeNameCounter <= 50; AttributeNameCounter++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrAttributeNameColNo].StringValue.Trim().Length > 0)  //only if attribute name exists
                            NoOfAttributeNames++;

                        if (wsErrorReport.Cells[cdrRowNo, cdrAttributeValueColNo].StringValue.Trim().Length > 0)  //only if attribute value exists
                            NoOfAttributeValuesFilled++;

                        cdrAttributeNameColNo = cdrAttributeNameColNo + 2;
                        cdrAttributeValueColNo = cdrAttributeValueColNo + 2;
                    }

                    if (NoOfAttributeNames > 0)
                        wsErrorReport.Cells[cdrRowNo, cdrNoOfAttributeNamesColNo].PutValue(NoOfAttributeNames);

                    if (NoOfAttributeValuesFilled > 0)
                        wsErrorReport.Cells[cdrRowNo, cdrNoOfAttributeValuesFilledColNo].PutValue(NoOfAttributeValuesFilled);
                }
                #endregion

                #region Writing Total Cell Count which includes Number of Attributes Names + Noun & Modifier + UNSPSC + URL 
                int cdrTotalCellCountColNo = wsErrorReport.Cells.MaxColumn + 1;

                if (IsNounExist)
                    wsErrorReport.Cells[0, cdrTotalCellCountColNo].PutValue("Total Cell Count - Includes -- Number of Attributes + Noun");

                if (IsModifierExist)
                    wsErrorReport.Cells[0, cdrTotalCellCountColNo].PutValue(wsErrorReport.Cells[0, cdrTotalCellCountColNo].StringValue + " + Modifier");

                if (IsUNSPSCExist)
                    wsErrorReport.Cells[0, cdrTotalCellCountColNo].PutValue(wsErrorReport.Cells[0, cdrTotalCellCountColNo].StringValue + " + UNSPSC");

                if (IsRefURL1Exist)
                    wsErrorReport.Cells[0, cdrTotalCellCountColNo].PutValue(wsErrorReport.Cells[0, cdrTotalCellCountColNo].StringValue + " + URL");

                int TotalCellCount = 0;
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    TotalCellCount = 0;
                    if (wsErrorReport.Cells[cdrRowNo, cdrNoOfAttributeNamesColNo].StringValue.Trim().Length > 0)
                        TotalCellCount = Convert.ToInt32(wsErrorReport.Cells[cdrRowNo, cdrNoOfAttributeNamesColNo].StringValue.Trim());

                    if (IsNounExist)
                        TotalCellCount = TotalCellCount + 1;

                    if (IsModifierExist)
                        TotalCellCount = TotalCellCount + 1;

                    if (IsUNSPSCExist)
                        TotalCellCount = TotalCellCount + 1;

                    if (IsRefURL1Exist)
                        TotalCellCount = TotalCellCount + 1;

                    wsErrorReport.Cells[cdrRowNo, cdrTotalCellCountColNo].PutValue(TotalCellCount);
                }
                #endregion

                #region Writing Total Errors Per Row
                int cdrTotalErrorsPerRowColNo = wsErrorReport.Cells.MaxColumn + 1;
                wsErrorReport.Cells[0, cdrTotalErrorsPerRowColNo].PutValue("Total Errors Per Row (Excluded URL Changed)");

                int TotalErrorsPerRow = 0;
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    TotalErrorsPerRow = 0;
                    if (IsRefURL1Exist)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrReferenceURL1MatchRemarksColNo].StringValue.Trim().ToUpper() == "PRODUCTION URL REMOVED")
                            TotalErrorsPerRow++;
                    }

                    if (IsUNSPSCExist)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrIsUNSPSCMatchedColNo].StringValue.Trim().Length > 0)
                            TotalErrorsPerRow++;
                    }

                    if (IsNounExist)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrIsNounMatchedColNo].StringValue.Trim().Length > 0)
                            TotalErrorsPerRow++;
                    }

                    if (IsModifierExist)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrIsModifierMatchedColNo].StringValue.Trim().Length > 0)
                            TotalErrorsPerRow++;
                    }

                    if (wsErrorReport.Cells[cdrRowNo, cdrHowManyAttributeValuesNotMatchedColNo].StringValue.Trim().Length > 0)
                        TotalErrorsPerRow++;

                    wsErrorReport.Cells[cdrRowNo, cdrTotalErrorsPerRowColNo].PutValue(TotalErrorsPerRow);
                }
                #endregion

                #region Writing "Attribute Value Not Matched - Error Rows Counter"
                int cdrAttributeValueNotMatchedErrorRowsCounterColNo = wsErrorReport.Cells.MaxColumn + 1;
                wsErrorReport.Cells[0, cdrAttributeValueNotMatchedErrorRowsCounterColNo].PutValue("Attribute Value Not Matched - Error Rows Counter");
                int ErrorRowNoCounter = 0;
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    if (wsErrorReport.Cells[cdrRowNo, cdrHowManyAttributeValuesNotMatchedColNo].StringValue.Length > 0)
                    {
                        if (Convert.ToInt32(wsErrorReport.Cells[cdrRowNo, cdrHowManyAttributeValuesNotMatchedColNo].StringValue) > 0)
                        {
                            ErrorRowNoCounter++;
                            wsErrorReport.Cells[cdrRowNo, cdrAttributeValueNotMatchedErrorRowsCounterColNo].PutValue(ErrorRowNoCounter);
                        }
                    }
                }
                #endregion

                #region Writing Summary Worksheet Report
                #region Setting Headers style
                Aspose.Cells.Style styleYellowColor = wsSummaryReport.Cells[1, 0].GetStyle();
                Aspose.Cells.Style styleBorders = wsSummaryReport.Cells[1, 0].GetStyle();

                //Setting Yellow fill-color style and borders
                styleYellowColor.ForegroundColor = System.Drawing.Color.FromArgb(0, 255, 255, 0);
                styleYellowColor.BackgroundColor = System.Drawing.Color.FromArgb(0, 255, 255, 0);
                styleYellowColor.Pattern = BackgroundType.VerticalStripe;
                styleYellowColor.Font.Color = System.Drawing.Color.Black;
                styleYellowColor.HorizontalAlignment = TextAlignmentType.Center;
                styleYellowColor.VerticalAlignment = TextAlignmentType.Center;
                styleYellowColor.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleYellowColor.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleYellowColor.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleYellowColor.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleYellowColor.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleYellowColor.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleYellowColor.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleYellowColor.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                //Setting borders for data rows
                styleBorders.HorizontalAlignment = TextAlignmentType.Center;
                styleBorders.VerticalAlignment = TextAlignmentType.Center;
                styleBorders.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleBorders.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleBorders.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleBorders.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleBorders.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleBorders.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleBorders.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleBorders.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                #endregion

                //Writing headers
                wsSummaryReport.Cells[0, 0].PutValue("Information");
                wsSummaryReport.Cells[0, 1].PutValue("Count");
                wsSummaryReport.Cells[0, 2].PutValue("Percentage");
                for (int c = 0; c <= 2; c++)
                    wsSummaryReport.Cells[0, c].SetStyle(styleYellowColor);

                //Writing Total Number of Records
                int TotalNoOfDataRowsInCDRWorksheet = wsErrorReport.Cells.MaxRow;
                wsSummaryReport.Cells[1, 0].PutValue("Total No. of Records");
                wsSummaryReport.Cells[1, 1].PutValue(TotalNoOfDataRowsInCDRWorksheet);
                wsSummaryReport.Cells[1, 2].PutValue("100%");

                int SummaryWorksheetRowNo = 1;
                //Writing UNSPSC Changed
                int CountOfIsUNSPSCMatched = 0;
                decimal UNSPSCChangedPercentage = 0;
                if (IsUNSPSCExist)
                {
                    SummaryWorksheetRowNo++;
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 0].PutValue("UNSPSC Changed");
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrIsUNSPSCMatchedColNo].StringValue.Trim().Length > 0)
                            CountOfIsUNSPSCMatched++;
                    }
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 1].PutValue(CountOfIsUNSPSCMatched);
                    UNSPSCChangedPercentage = Convert.ToDecimal((CountOfIsUNSPSCMatched * 100.00) / TotalNoOfDataRowsInCDRWorksheet);
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 2].PutValue(UNSPSCChangedPercentage.ToString("0.00") + "%");
                }

                //Writing URL Added / Removed Count
                int CountOfRefURL1MatchRemarks = 0;
                decimal URLAddedORRemovedPercentage = 0;
                if (IsRefURL1Exist)
                {
                    SummaryWorksheetRowNo++;
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 0].PutValue("URL Added / Removed");
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrReferenceURL1MatchRemarksColNo].StringValue.Trim().ToUpper() == "PRODUCTION URL REMOVED")
                            CountOfRefURL1MatchRemarks++;
                    }
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 1].PutValue(CountOfRefURL1MatchRemarks);
                    URLAddedORRemovedPercentage = Convert.ToDecimal((CountOfRefURL1MatchRemarks * 100.00) / TotalNoOfDataRowsInCDRWorksheet);
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 2].PutValue(URLAddedORRemovedPercentage.ToString("0.00") + "%");
                }

                //Writing Noun, Modifier Changed
                int CountOfNounModifierChanged = 0;
                decimal NounModifierChangedPercentage = 0;
                if (IsNounExist || IsModifierExist)
                {
                    SummaryWorksheetRowNo++;
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 0].PutValue("Noun / Modifier Changed");
                    for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                    {
                        if (wsErrorReport.Cells[cdrRowNo, cdrIsNounMatchedColNo].StringValue.Trim().Length > 0 ||
                            wsErrorReport.Cells[cdrRowNo, cdrIsModifierMatchedColNo].StringValue.Trim().Length > 0)
                            CountOfNounModifierChanged++;
                    }
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 1].PutValue(CountOfNounModifierChanged);
                    NounModifierChangedPercentage = Convert.ToDecimal((CountOfNounModifierChanged * 100.00) / TotalNoOfDataRowsInCDRWorksheet);
                    wsSummaryReport.Cells[SummaryWorksheetRowNo, 2].PutValue(NounModifierChangedPercentage.ToString("0.00") + "%");
                }

                //Writing Attribute Values Changed (Excluding Noun & Modifer changed lines)
                int CountOfAttributeValuesNotMatched = 0;
                decimal AttributeValuesNotMatchedPercentage = 0;
                SummaryWorksheetRowNo++;
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 0].PutValue("Attribute Values Changed Lines (Excluding Noun & Modifer changed lines)");
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    if (wsErrorReport.Cells[cdrRowNo, cdrHowManyAttributeValuesNotMatchedColNo].StringValue.Trim().Length > 0)
                    {
                        if (Convert.ToInt32(wsErrorReport.Cells[cdrRowNo, cdrHowManyAttributeValuesNotMatchedColNo].StringValue.Trim()) > 0)
                            CountOfAttributeValuesNotMatched++;
                    }
                }
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 1].PutValue(CountOfAttributeValuesNotMatched);
                AttributeValuesNotMatchedPercentage = Convert.ToDecimal((CountOfAttributeValuesNotMatched * 100.00) / TotalNoOfDataRowsInCDRWorksheet);
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 2].PutValue(AttributeValuesNotMatchedPercentage.ToString("0.00") + "%");

                //Writing DPO - Defects Per Opportunity
                int SumOfTotalCellCount = 0, SumOfTotalErrorPerRow = 0;
                decimal DPOPercentage = 0;
                SummaryWorksheetRowNo++;
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 0].PutValue("Defects Per Opportunity (Sum of Total Errors Per Row / Sum of Total Cell Count)");
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    if (wsErrorReport.Cells[cdrRowNo, cdrTotalCellCountColNo].StringValue.Trim().Length > 0)
                        SumOfTotalCellCount = SumOfTotalCellCount + Convert.ToInt32(wsErrorReport.Cells[cdrRowNo, cdrTotalCellCountColNo].StringValue.Trim());

                    if (wsErrorReport.Cells[cdrRowNo, cdrTotalErrorsPerRowColNo].StringValue.Trim().Length > 0)
                        SumOfTotalErrorPerRow = SumOfTotalErrorPerRow + Convert.ToInt32(wsErrorReport.Cells[cdrRowNo, cdrTotalErrorsPerRowColNo].StringValue.Trim());
                }
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 1].PutValue(SumOfTotalErrorPerRow + "/" + SumOfTotalCellCount);
                DPOPercentage = Convert.ToDecimal((SumOfTotalErrorPerRow * 100.00) / SumOfTotalCellCount);
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 2].PutValue(DPOPercentage.ToString("0.00") + "%");

                //Writing Total lines affected with changes
                int CountOfLinesAffectedWithChanges = 0;
                decimal LinesAffectedWithChangesPercentage = 0;
                SummaryWorksheetRowNo++;
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 0].PutValue("Total lines affected with changes");
                for (int cdrRowNo = 1; cdrRowNo <= wsErrorReport.Cells.MaxRow; cdrRowNo++)
                {
                    if (wsErrorReport.Cells[cdrRowNo, cdrReferenceURL1MatchRemarksColNo].StringValue.Trim().ToUpper() == "PRODUCTION URL REMOVED" ||
                            wsErrorReport.Cells[cdrRowNo, cdrIsUNSPSCMatchedColNo].StringValue.Trim().Length > 0 ||
                            wsErrorReport.Cells[cdrRowNo, cdrIsNounMatchedColNo].StringValue.Trim().Length > 0 ||
                            wsErrorReport.Cells[cdrRowNo, cdrIsModifierMatchedColNo].StringValue.Trim().Length > 0 ||
                            wsErrorReport.Cells[cdrRowNo, cdrHowManyAttributeValuesNotMatchedColNo].StringValue.Trim().Length > 0
                       )
                        CountOfLinesAffectedWithChanges++;
                }
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 1].PutValue(CountOfLinesAffectedWithChanges);
                LinesAffectedWithChangesPercentage = Convert.ToDecimal((CountOfLinesAffectedWithChanges * 100.00) / TotalNoOfDataRowsInCDRWorksheet);
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 2].PutValue(LinesAffectedWithChangesPercentage.ToString("0.00") + "%");

                //Writing First Pass Yield
                int TotalLinesWithoutChanges = TotalNoOfDataRowsInCDRWorksheet - CountOfLinesAffectedWithChanges;
                decimal FPYPercentage = 0;
                FPYPercentage = Convert.ToDecimal((TotalLinesWithoutChanges * 100.00) / TotalNoOfDataRowsInCDRWorksheet);
                SummaryWorksheetRowNo++;
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 0].PutValue("First Pass Yield");
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 1].PutValue(TotalLinesWithoutChanges);
                wsSummaryReport.Cells[SummaryWorksheetRowNo, 2].PutValue(FPYPercentage.ToString("0.00") + "%");

                //Setting border style for all data rows
                for (int r = 1; r <= wsSummaryReport.Cells.MaxRow; r++)
                {
                    for (int c = 0; c <= 2; c++)
                        wsSummaryReport.Cells[r, c].SetStyle(styleBorders);
                }

                wsSummaryReport.AutoFitColumns();
                #endregion

                #region Inserting columns from Prod completed worksheet in Error Report worksheet for ease of checking
                //Worksheet source = wsProduction;
                Worksheet destination = wsErrorReport;

                #region Reference URL 1
                Aspose.Cells.Style styleRefURL1 = wsProduction.Cells[0, pcReferenceURL1ColNo].GetStyle();

                CellArea ca = new CellArea();
                ca.StartRow = 0;
                ca.EndRow = wsProduction.Cells.MaxRow;
                ca.StartColumn = pcReferenceURL1ColNo;
                ca.EndColumn = pcReferenceURL1ColNo;

                destination.Cells.InsertRange(ca, 1, ShiftType.Right);
                //destination.Cells.InsertColumn(pcReferenceURL1ColNo);         /*InsertColumn method is not working, so alternate InsertRange method is used*/
                destination.Cells.ClearFormats(1, pcReferenceURL1ColNo, destination.Cells.MaxRow, pcReferenceURL1ColNo);

                //Copying values from Production completed worksheet to Error Report worksheet
                destination.Cells[0, pcReferenceURL1ColNo].PutValue("Reference URL 1 (from Prod Completed)");
                destination.Cells[0, pcReferenceURL1ColNo].SetStyle(styleRefURL1);
                for (int r = 1; r <= wsProduction.Cells.MaxRow; r++)
                    destination.Cells[r, pcReferenceURL1ColNo].PutValue(wsProduction.Cells[r, pcReferenceURL1ColNo].StringValue);
                #endregion

                #region UNSPSC Code and Description
                Aspose.Cells.Style styleUNSPSC = wsProduction.Cells[0, pcUNSPSCCodeColNo].GetStyle();
                ca.StartRow = 0;
                ca.EndRow = wsProduction.Cells.MaxRow;
                ca.StartColumn = pcUNSPSCCodeColNo + 1;
                ca.EndColumn = pcUNSPSCCodeColNo + 1;
                destination.Cells.InsertRange(ca, 1, ShiftType.Right);

                ca.StartRow = 0;
                ca.EndRow = wsProduction.Cells.MaxRow;
                ca.StartColumn = pcUNSPSCDescriptionColNo + 1;
                ca.EndColumn = pcUNSPSCDescriptionColNo + 1;
                destination.Cells.InsertRange(ca, 1, ShiftType.Right);

                destination.Cells.ClearFormats(1, pcUNSPSCCodeColNo + 1, destination.Cells.MaxRow, pcUNSPSCCodeColNo + 1);
                destination.Cells.ClearFormats(1, pcUNSPSCDescriptionColNo + 1, destination.Cells.MaxRow, pcUNSPSCDescriptionColNo + 1);

                destination.Cells[0, pcUNSPSCCodeColNo + 1].PutValue("UNSPSC Code (from Prod Completed)");
                destination.Cells[0, pcUNSPSCDescriptionColNo + 1].PutValue("UNSPSC Description (from Prod Completed)");

                destination.Cells[0, pcUNSPSCCodeColNo + 1].SetStyle(styleUNSPSC);
                destination.Cells[0, pcUNSPSCDescriptionColNo + 1].SetStyle(styleUNSPSC);
                for (int r = 1; r <= wsProduction.Cells.MaxRow; r++)
                {
                    destination.Cells[r, pcUNSPSCCodeColNo + 1].PutValue(wsProduction.Cells[r, pcUNSPSCCodeColNo].StringValue);
                    destination.Cells[r, pcUNSPSCDescriptionColNo + 1].PutValue(wsProduction.Cells[r, pcUNSPSCDescriptionColNo].StringValue);
                }
                #endregion

                #region Noun
                Aspose.Cells.Style styleNoun = wsProduction.Cells[0, pcNounColNo].GetStyle();
                ca.StartRow = 0;
                ca.EndRow = wsProduction.Cells.MaxRow;
                ca.StartColumn = pcNounColNo + 3;
                ca.EndColumn = pcNounColNo + 3;
                destination.Cells.InsertRange(ca, 1, ShiftType.Right);
                //destination.Cells.InsertColumn(pcNounColNo + 3);
                destination.Cells.ClearFormats(1, pcNounColNo + 3, destination.Cells.MaxRow, pcNounColNo + 3);

                destination.Cells[0, pcNounColNo + 3].PutValue("Noun (from Prod Completed)");

                destination.Cells[0, pcNounColNo + 3].SetStyle(styleNoun);
                for (int r = 1; r <= wsProduction.Cells.MaxRow; r++)
                    destination.Cells[r, pcNounColNo + 3].PutValue(wsProduction.Cells[r, pcNounColNo].StringValue);
                #endregion

                #region Modifier
                Aspose.Cells.Style styleModifier = wsProduction.Cells[0, pcModifierColNo].GetStyle();
                ca.StartRow = 0;
                ca.EndRow = wsProduction.Cells.MaxRow;
                ca.StartColumn = pcModifierColNo + 4;
                ca.EndColumn = pcModifierColNo + 4;
                destination.Cells.InsertRange(ca, 1, ShiftType.Right);
                destination.Cells.ClearFormats(1, pcModifierColNo + 4, destination.Cells.MaxRow, pcModifierColNo + 4);

                destination.Cells[0, pcModifierColNo + 4].PutValue("Modifier (from Prod Completed)");
                destination.Cells[0, pcModifierColNo + 4].SetStyle(styleModifier);
                for (int r = 1; r <= wsProduction.Cells.MaxRow; r++)
                    destination.Cells[r, pcModifierColNo + 4].PutValue(wsProduction.Cells[r, pcModifierColNo].StringValue);
                #endregion

                #region Attribute Value 1 to 50
                int rAttributeValueColNo = cdrAttributeValue1ColNo + 5;
                Aspose.Cells.Style styleAttributeValue = wsProduction.Cells[0, pcAttributeValue1ColNo].GetStyle();
                for (int aValueCounter = 1; aValueCounter <= 50; aValueCounter++)
                {
                    ca.StartRow = 0;
                    ca.EndRow = wsProduction.Cells.MaxRow;
                    ca.StartColumn = rAttributeValueColNo;
                    ca.EndColumn = rAttributeValueColNo;
                    destination.Cells.InsertRange(ca, 1, ShiftType.Right);
                    //destination.Cells.InsertColumn(rAttributeValueColNo);
                    destination.Cells.ClearFormats(1, rAttributeValueColNo, destination.Cells.MaxRow, rAttributeValueColNo);

                    destination.Cells[0, rAttributeValueColNo].PutValue("Attribute Value " + aValueCounter.ToString() + " (from Prod Completed)");
                    destination.Cells[0, rAttributeValueColNo].SetStyle(styleAttributeValue);
                    for (int r = 1; r <= wsProduction.Cells.MaxRow; r++)
                        destination.Cells[r, rAttributeValueColNo].PutValue(wsProduction.Cells[r, pcAttributeValue1ColNo].StringValue);
                    pcAttributeValue1ColNo = pcAttributeValue1ColNo + 2;
                    rAttributeValueColNo = rAttributeValueColNo + 3;
                }
                #endregion
                #endregion

                #region Save and Download the file
                wsErrorReport.AutoFitColumns();
                wsSummaryReport.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName.Trim();
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
