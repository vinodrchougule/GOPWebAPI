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
    [RoutePrefix("api/FindSameDrugs")]
    public class FindSameDrugsController : ApiController
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
                if (ws1.Cells[0, 0].StringValue.Trim().ToUpper() != "GENERIC DESC CODE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [A1] with Value 'Generic Desc Code' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells[0, 1].StringValue.Trim().ToUpper() != "GENERIC DESC")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Cell [B1] with Value 'Generic Desc' not found in input file first worksheet. Please select a valid file.");
                }

                if (ws1.Cells.Rows.Count <= 1)
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

        #region Process And Write the output to file
        [HttpPost]
        [Route("FindSameDrugsAndWriteToOutput")]
        public HttpResponseMessage FindSameDrugsAndWriteToOutput(string UploadedInputFileName, string InputFileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);

                Worksheet wsGD = wbIF.Worksheets[0];          //Generic Description
                int maxGDRows = wsGD.Cells.MaxRow;

                #region Writing Headers
                wsGD.Cells[0, 2].PutValue("Value1");
                wsGD.Cells[0, 3].PutValue("Value2");
                wsGD.Cells[0, 4].PutValue("Value3");
                wsGD.Cells[0, 5].PutValue("Value4");
                wsGD.Cells[0, 6].PutValue("Value5");
                wsGD.Cells[0, 7].PutValue("Duplicate Counter");
                wsGD.Cells[0, 8].PutValue("Final Column");
                #endregion

                int ctr = 1;
                string genericMainDescription = string.Empty, genericIteratingDescription = string.Empty, Value1 = string.Empty;
                bool AreAllWordsMatching = false, IsCounterUsed = false;

                #region Write the values and duplicate counter
                for (int gdmr = 1; gdmr < maxGDRows; gdmr++)       //Generic Description Main Row
                {
                    AreAllWordsMatching = false; IsCounterUsed = false;
                    genericMainDescription = wsGD.Cells[gdmr, 1].StringValue.Trim();
                    genericMainDescription = genericMainDescription.Replace(",", "/");
                    genericMainDescription = genericMainDescription.Replace('+', '/');
                    genericMainDescription = ReplaceMGwithBlank(genericMainDescription);
                    string[] genericMainDescriptionArray = genericMainDescription.Split('/');
                    Value1 = wsGD.Cells[gdmr, 2].StringValue.Trim();

                    if (string.IsNullOrEmpty(Value1) && genericMainDescriptionArray.Length > 0)         //&& genericMainDescriptionArray.Length <= 5
                    {
                        for (int gdir = gdmr + 1; gdir <= maxGDRows; gdir++)   //Generic Description Iterating Row
                        {
                            if (string.IsNullOrEmpty(wsGD.Cells[gdir, 2].StringValue.Trim()))       // if iterating row Value1 is null or empty
                            {
                                genericIteratingDescription = wsGD.Cells[gdir, 1].StringValue.Trim();
                                genericIteratingDescription = genericIteratingDescription.Replace(", ", "/");
                                genericIteratingDescription = genericIteratingDescription.Replace('+', '/');
                                genericIteratingDescription = ReplaceMGwithBlank(genericIteratingDescription);
                                string[] genericIteratingDescriptionArray = genericIteratingDescription.Split('/');
                                if (genericIteratingDescriptionArray.Length > 0)            //&& genericIteratingDescriptionArray.Length <= 5
                                {
                                    if (genericMainDescriptionArray.Length == genericIteratingDescriptionArray.Length)
                                    {
                                        AreAllWordsMatching = false;
                                        for (int mi = 0; mi < genericMainDescriptionArray.Length; mi++)        //Main Index
                                        {
                                            if (!string.IsNullOrEmpty(genericMainDescriptionArray[mi].Trim()))
                                            {
                                                AreAllWordsMatching = false;
                                                for (int ii = 0; ii < genericIteratingDescriptionArray.Length; ii++)        //Iterating Index
                                                {
                                                    if (genericMainDescriptionArray[mi].Trim().ToUpper() == genericIteratingDescriptionArray[ii].Trim().ToUpper())
                                                    {
                                                        AreAllWordsMatching = true;
                                                        break;
                                                    }
                                                }

                                                if (!AreAllWordsMatching)
                                                    break;
                                            }
                                        }

                                        if (AreAllWordsMatching)
                                        {
                                            for (int mi = 0; mi < genericMainDescriptionArray.Length; mi++)        //Main Index
                                            {
                                                wsGD.Cells[gdmr, mi + 2].PutValue(genericMainDescriptionArray[mi].Trim());
                                                wsGD.Cells[gdir, mi + 2].PutValue(genericMainDescriptionArray[mi].Trim());
                                                if (mi > 4)
                                                    break;
                                            }
                                            wsGD.Cells[gdmr, 7].PutValue(ctr);
                                            wsGD.Cells[gdir, 7].PutValue(ctr);
                                            IsCounterUsed = true;
                                        }
                                    }
                                }
                            }
                        }
                        if (IsCounterUsed)
                            ctr++;
                    }
                }
                #endregion

                #region Writing Non-duplicate values from 1 to 5 and Final Column Values
                string genericDescription = string.Empty;
                string FinalColumnValue = string.Empty;
                for (int gdr = 1; gdr <= maxGDRows; gdr++)
                {
                    if (string.IsNullOrEmpty(wsGD.Cells[gdr, 7].StringValue.Trim()))
                    {
                        genericDescription = wsGD.Cells[gdr, 1].StringValue.Trim();
                        genericDescription = genericDescription.Replace(",", "+");
                        genericDescription = genericDescription.Replace('/', '+');
                        string[] genericDescriptionArray = genericDescription.Split('+');
                        if (genericDescriptionArray.Length > 0)        //&& genericDescriptionArray.Length<=5
                        {
                            for (int i = 0; i < genericDescriptionArray.Length; i++)
                            {
                                if (i == 0)
                                    wsGD.Cells[gdr, 2].PutValue(genericDescriptionArray[i].Trim());
                                else if (i == 1)
                                    wsGD.Cells[gdr, 3].PutValue(genericDescriptionArray[i].Trim());
                                else if (i == 2)
                                    wsGD.Cells[gdr, 4].PutValue(genericDescriptionArray[i].Trim());
                                else if (i == 3)
                                    wsGD.Cells[gdr, 5].PutValue(genericDescriptionArray[i].Trim());
                                else if (i == 4)
                                    wsGD.Cells[gdr, 6].PutValue(genericDescriptionArray[i].Trim());
                                else
                                    break;
                            }
                        }
                    }
                    FinalColumnValue = string.Empty;
                    for (int vi = 2; vi <= 6; vi++)
                    {
                        if (!string.IsNullOrEmpty(wsGD.Cells[gdr, vi].StringValue.Trim()))
                        {
                            if (FinalColumnValue.Trim().Length > 0)
                                FinalColumnValue = FinalColumnValue + " + " + wsGD.Cells[gdr, vi].StringValue.Trim();
                            else
                                FinalColumnValue = wsGD.Cells[gdr, vi].StringValue.Trim();
                        }
                    }
                    wsGD.Cells[gdr, 8].PutValue(FinalColumnValue);
                }
                #endregion

                #region Write Unique List of Drug Names in second worksheet
                bool IsDrugNameExists = false;
                if (wbIF.Worksheets.Count == 1)
                    wbIF.Worksheets.Add("Unique List of Drug Names");
                else
                    wbIF.Worksheets[1].Name = "Unique List of Drug Names";
                Worksheet wsUD = wbIF.Worksheets[1];
                wsUD.Cells[0, 0].PutValue("Unique Drugs");
                for (int GDRow = 1; GDRow <= wsGD.Cells.MaxRow; GDRow++)
                {
                    genericDescription = wsGD.Cells[GDRow, 1].StringValue.Trim();
                    genericDescription = genericDescription.Replace(",", "/");
                    genericDescription = genericDescription.Replace('+', '/');
                    genericDescription = Regex.Replace(genericDescription, "[.0-9]", "").Trim();
                    genericDescription = genericDescription.ToUpper().Replace(" MG ", "");
                    genericDescription = genericDescription.ToUpper().Replace(" MG", "");
                    genericDescription = genericDescription.ToUpper().Replace("%", "");
                    genericDescription = genericDescription.ToUpper().Replace("-", "");
                    string[] genericDescriptionArray = genericDescription.Split('/');
                    for (int dni = 0; dni < genericDescriptionArray.Length; dni++)
                    {
                        IsDrugNameExists = false;
                        for (int udr = 1; udr <= wsUD.Cells.MaxRow; udr++)
                        {
                            if (wsUD.Cells[udr, 0].StringValue.Trim().ToUpper() == genericDescriptionArray[dni].Trim().ToUpper())
                            {
                                IsDrugNameExists = true;
                                break;
                            }
                        }

                        if (genericDescriptionArray[dni].Trim().Length > 0)
                        {
                            if (!IsDrugNameExists)
                                wsUD.Cells[wsUD.Cells.MaxRow + 1, 0].PutValue(genericDescriptionArray[dni].Trim());
                        }
                    }
                }
                #endregion

                #region Save the file
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wbIF.Save(filename);
                #endregion

                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Helping Functions
        private string ReplaceMGwithBlank(string genericDescription)
        {
            int mgIndex = -1;

            if (genericDescription.ToUpper().IndexOf("MG") > 0)
            {
                while (genericDescription.ToUpper().Contains("MG"))
                {
                    mgIndex = genericDescription.ToUpper().IndexOf("MG");
                    if (genericDescription[mgIndex - 1] == ' ')
                        mgIndex = mgIndex - 2;
                    while (mgIndex > 0)
                    {
                        if (Char.IsDigit(genericDescription[mgIndex]) || genericDescription[mgIndex] == '.')
                            genericDescription = genericDescription.Remove(mgIndex, 1);

                        mgIndex--;
                        if (genericDescription[mgIndex] == ' ')
                            break;
                    }
                    genericDescription = genericDescription.Remove(genericDescription.ToUpper().IndexOf("MG"), 2);
                }
            }

            return genericDescription;
        }
        #endregion
    }
}
