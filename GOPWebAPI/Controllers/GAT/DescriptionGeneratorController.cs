using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Web;
using System.Web.Http;
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;
using Newtonsoft.Json.Linq;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/DescriptionGenerator")]
    public class DescriptionGeneratorController : ApiController
    {
        private BLLGATDescriptionGenerator _BLLGATDescriptionGenerator;

        public DescriptionGeneratorController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGATDescriptionGenerator = new BLLGATDescriptionGenerator(connectionString);
        }

        #region Read Description Generator Setting Names
        [HttpGet]
        [Route("ReadDGSettingNames")]
        public IHttpActionResult ReadDGSettingNames()
        {
            try
            {
                DataTable dataTable = _BLLGATDescriptionGenerator.ReadDGSettingNames();
                List<string> DGSettingNamesList = new List<string>();

                foreach (DataRow row in dataTable.Rows)
                    DGSettingNamesList.Add(row["SettingName"].ToString());

                return Ok(DGSettingNamesList);
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Validate Input File Data
        [HttpPost]
        [Route("ValidateInputFileData")]
        public HttpResponseMessage ValidateInputFileData(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "REFERENCE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Reference' not found in input file first worksheet. Please select a valid file.");
                }

                #region Validating MFR Name, MFR Part No. columns
                string hdr;
                int MFRNo = 1;
                for (int mCtr = 1; mCtr <= 20; mCtr += 2)
                {
                    hdr = wsIF.Cells[0, mCtr].StringValue.Trim().ToUpper();

                    if (hdr != "MFR NAME" + MFRNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First row and column no. " + mCtr.ToString() + " with heading 'MFR Name'" + MFRNo.ToString() + " not found in input file first worksheet. Please select a valid file.");
                    }

                    hdr = wsIF.Cells[0, mCtr + 1].StringValue.Trim().ToUpper();
                    if (hdr != "MFR PART NO" + MFRNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column no. " + (mCtr + 1).ToString() + " with heading 'MFR Part No'" + MFRNo.ToString() + " not found in input file first worksheet. Please select a valid file.");
                    }
                    MFRNo++;
                }
                #endregion

                if (wsIF.Cells[0, 21].StringValue.Trim().ToUpper() != "NOUN")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column no.22 with heading 'Noun'  not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 22].StringValue.Trim().ToUpper() != "MODIFIER")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column no.23 with heading 'Modifier' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 23].StringValue.Trim().ToUpper() != "ADDITIONAL INFO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column no.24 with heading 'Additional Info' not found in input file first worksheet. Please select a valid file.");
                }

                #region Validating Attribute columns
                int attNo = 1;
                for (int aCtr = 1; aCtr <= 100; aCtr += 2)
                {
                    hdr = wsIF.Cells[0, aCtr + 23].StringValue.Trim().ToUpper();

                    if (hdr != "ATTRIBUTE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column No. " + (aCtr + 23).ToString() + " must contain 'Attribute" + attNo.ToString() + "' heading.Please download format from Template.");
                    }

                    hdr = wsIF.Cells[0, aCtr + 24].StringValue.Trim().ToUpper();
                    if (hdr != "VALUE" + attNo.ToString())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Column No. " + (aCtr + 24).ToString() + " must contain 'Value" + attNo.ToString() + "' heading.Please download format from Template.");
                    }
                    attNo++;
                }
                #endregion

                if (wsIF.Cells[0, 124].StringValue.Trim().ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "One hundred Twenty Fourth Column with heading 'Description' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 125].StringValue.Trim().ToUpper() != "LEFT-OUT INFORMATION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "One hundred Twenty Fifth Column with heading 'Left-Out Information' not found in input file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Input file data validated Successfully.");
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

        #region Validate Abbreviation File Data
        [HttpPost]
        [Route("ValidateAbbreviationFileData")]
        public HttpResponseMessage ValidateAbbreviationFileData(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                Workbook wbAF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbAF.Open(UploadedFilepath);
                Worksheet wsAF = wbAF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template
                if (wsAF.Cells[0, 0].StringValue.Trim().ToUpper() != "EXPANDED VERSION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Expanded Version' not found in input abbreviation file first worksheet. Please select a valid file.");
                }

                if (wsAF.Cells[0, 1].StringValue.Trim().ToUpper() != "ABBREVIATION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Abbreviation' not found in input abbreviation file first worksheet. Please select a valid file.");
                }

                if (wsAF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Abbreviation File first Worksheet has no data rows.");
                }
                #endregion

                #region Expanded Version and Abbreviation should not be same
                for (int row = 1; row <= wsAF.Cells.MaxRow; row++)
                {
                    if (wsAF.Cells[row, 0].StringValue.Trim().ToUpper() == wsAF.Cells[row, 1].StringValue.Trim().ToUpper())
                    {
                        File.Delete(UploadedFilepath);
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Expanded version and Abbreviation both cannot be same. Expanded Version: " + wsAF.Cells[row, 0].StringValue + " ,Row No:" + row.ToString());
                    }
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Abbreviation file data validated Successfully.");
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

        #region Validate Identifiers File Data
        [HttpPost]
        [Route("ValidateIdentifiersFileData")]
        public HttpResponseMessage ValidateIdentifiersFileData(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (!File.Exists(UploadedFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "File not found.");

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Check uploaded file first worksheet has columns as per template

                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "CHARACTERISTICS")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'Characteristics' not found in identifiers file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 1].StringValue.Trim().ToUpper() != "ABBREVIATION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'Abbreviation' not found in identifiers file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells[0, 2].StringValue.Trim().ToUpper() != "SUFFIX OR PREFIX")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Suffix OR Prefix' not found in identifiers file first worksheet. Please select a valid file.");
                }

                if (wsIF.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Identifiers File first Worksheet has no data rows.");
                }
                #endregion

                return Request.CreateResponse(HttpStatusCode.OK, "Identifiers file data validated Successfully.");
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

        #region Save Description Generator Settings
        [HttpPost]
        [Route("SaveDescriptionGeneratorSettings")]
        public HttpResponseMessage SaveDescriptionGeneratorSettings([FromBody] DescriptionGenerator model)
        {
            try
            {
                if (string.IsNullOrEmpty(model.SettingName))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Setting Name is mandatory.");

                string Result = _BLLGATDescriptionGenerator.SaveDescriptionGeneratorSettings(model);

                if (!string.IsNullOrEmpty(Result) && Result.Trim().ToLower().StartsWith("success"))
                {
                    string[] strResult = Result.Split(',');

                    DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                    #region Saving Abbreviation and Identifier Files
                    string DGSetID = strResult[1];
                    if (!string.IsNullOrEmpty(model.AbbreviationFileName))
                    {
                        string NewAbbreviationFileName = DGSetID + " ~ " + model.AbbreviationFileName;
                        if (File.Exists(dirTemp + model.AbbreviationFileName))
                        {
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GAT/Description Generator/"));
                            FileOperations.MoveFile(dirTemp, model.AbbreviationFileName, dirUploads, NewAbbreviationFileName);
                        }
                    }

                    if (model.IsToApplyIdentifiers)
                    {
                        string NewIdentifierFileName = DGSetID + " ~ " + model.IdentifierFileName;
                        if (File.Exists(dirTemp + model.IdentifierFileName))
                        {
                            DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GAT/Description Generator/"));
                            FileOperations.MoveFile(dirTemp, model.IdentifierFileName, dirUploads, NewIdentifierFileName);
                        }
                    }
                    #endregion

                    return Request.CreateResponse(HttpStatusCode.OK, "Settings saved Successfully.");
                }
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Description Generator Saved Setting
        [HttpGet]
        [Route("ReadDescriptionGeneratorSavedSetting")]
        public HttpResponseMessage ReadDescriptionGeneratorSavedSetting(string SettingName)
        {
            try
            {
                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                if (string.IsNullOrEmpty(SettingName) || string.IsNullOrEmpty(SettingName))
                {
                    Result objResult = new Result();
                    objResult.Msg = "Invalid Setting Name";
                    objResult.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.BadRequest, objResult);
                }

                DataTable dataTable = new DataTable();
                dataTable = _BLLGATDescriptionGenerator.ReadDescriptionGeneratorSavedSetting(SettingName);

                if (dataTable.Rows.Count > 0)
                {
                    JObject PEReport = new JObject(new JProperty("DGSettingID", dataTable.Rows[0]["DGSettingID"]),
                                                    new JProperty("IsNounExcluded", dataTable.Rows[0]["IsNounExcluded"]),
                                                    new JProperty("IsModifierExcluded", dataTable.Rows[0]["IsModifierExcluded"]),
                                                    new JProperty("IsAttributeNameExcluded", dataTable.Rows[0]["IsAttributeNameExcluded"]),
                                                    new JProperty("IsAttributeValueExcluded", dataTable.Rows[0]["IsAttributeValueExcluded"]),
                                                    new JProperty("IsAdditionalInformationExcluded", dataTable.Rows[0]["IsAdditionalInformationExcluded"]),
                                                    new JProperty("IsMFRNameExcluded", dataTable.Rows[0]["IsMFRNameExcluded"]),
                                                    new JProperty("IsMFRPartNoExcluded", dataTable.Rows[0]["IsMFRPartNoExcluded"]),
                                                    new JProperty("SpecificModifierExcluded", dataTable.Rows[0]["SpecificModifierExcluded"]),
                                                    new JProperty("IsToInterpretAdditionalInformation", dataTable.Rows[0]["IsToInterpretAdditionalInformation"]),
                                                    new JProperty("IsToIncludeAttributeNameFromAdditionalInformation", dataTable.Rows[0]["IsToIncludeAttributeNameFromAdditionalInformation"]),
                                                    new JProperty("IsToIncludeMaximumValues", dataTable.Rows[0]["IsToIncludeMaximumValues"]),
                                                    new JProperty("IsToInterpretAllAttributeValues", dataTable.Rows[0]["IsToInterpretAllAttributeValues"]),
                                                    new JProperty("IsNounToBeAbbreviated", dataTable.Rows[0]["IsNounToBeAbbreviated"]),
                                                    new JProperty("IsModifierToBeAbbreviated", dataTable.Rows[0]["IsModifierToBeAbbreviated"]),
                                                    new JProperty("IsAttributeNameToBeAbbreviated", dataTable.Rows[0]["IsAttributeNameToBeAbbreviated"]),
                                                    new JProperty("IsAttributeValueToBeAbbreviated", dataTable.Rows[0]["IsAttributeValueToBeAbbreviated"]),
                                                    new JProperty("IsAdditionalInformationToBeAbbreviated", dataTable.Rows[0]["IsAdditionalInformationToBeAbbreviated"]),
                                                    new JProperty("IsMFRNameToBeAbbreviated", dataTable.Rows[0]["IsMFRNameToBeAbbreviated"]),
                                                    new JProperty("AbbreviationFileName", dataTable.Rows[0]["AbbreviationFileName"]),
                                                    new JProperty("DelimiterAfterNoun", dataTable.Rows[0]["DelimiterAfterNoun"]),
                                                    new JProperty("DelimiterAfterModifier", dataTable.Rows[0]["DelimiterAfterModifier"]),
                                                    new JProperty("DelimiterAfterAttributeName", dataTable.Rows[0]["DelimiterAfterAttributeName"]),
                                                    new JProperty("DelimiterAfterAttributeValue", dataTable.Rows[0]["DelimiterAfterAttributeValue"]),
                                                    new JProperty("DelimiterAfterAdditionalInformation", dataTable.Rows[0]["DelimiterAfterAdditionalInformation"]),
                                                    new JProperty("DelimiterAfterMFRName", dataTable.Rows[0]["DelimiterAfterMFRName"]),
                                                    new JProperty("DelimiterAfterMFRPartNo", dataTable.Rows[0]["DelimiterAfterMFRPartNo"]),
                                                    new JProperty("MultipleValuesSeparator", dataTable.Rows[0]["MultipleValuesSeparator"]),
                                                    new JProperty("IsToApplyIdentifiers", dataTable.Rows[0]["IsToApplyIdentifiers"]),
                                                    new JProperty("IdentifierFileName", dataTable.Rows[0]["IdentifierFileName"]),
                                                    new JProperty("IsToAddSpaceBeforeORAfterIdentifier", dataTable.Rows[0]["IsToAddSpaceBeforeORAfterIdentifier"]),
                                                    new JProperty("IsToApplyIdentifierToAdditionalInformation", dataTable.Rows[0]["IsToApplyIdentifierToAdditionalInformation"]),
                                                    new JProperty("PrefixForAdditionalInformation", dataTable.Rows[0]["PrefixForAdditionalInformation"]),
                                                    new JProperty("PrefixForMFRName", dataTable.Rows[0]["PrefixForMFRName"]),
                                                    new JProperty("PrefixForMFRPartNo", dataTable.Rows[0]["PrefixForMFRPartNo"]),
                                                    new JProperty("IsToIncludeAttributeNameWithNULLValues", dataTable.Rows[0]["IsToIncludeAttributeNameWithNULLValues"]),
                                                    new JProperty("IsToIncludeAllOtherMFRNames", dataTable.Rows[0]["IsToIncludeAllOtherMFRNames"]),
                                                    new JProperty("IsToIncludeAllOtherMFRPartNos", dataTable.Rows[0]["IsToIncludeAllOtherMFRPartNos"]),
                                                    new JProperty("IsToPrefixAllMFRNames", dataTable.Rows[0]["IsToPrefixAllMFRNames"]),
                                                    new JProperty("IsToPrefixAllMFRPartNos", dataTable.Rows[0]["IsToPrefixAllMFRPartNos"]),
                                                    new JProperty("DescriptionToGenerate", dataTable.Rows[0]["DescriptionToGenerate"]),
                                                    new JProperty("TruncationType", dataTable.Rows[0]["TruncationType"]),
                                                    new JProperty("CharacterLimit", dataTable.Rows[0]["CharacterLimit"]),
                                                    new JProperty("DelimiterForTruncation", dataTable.Rows[0]["DelimiterForTruncation"]),
                                                    new JProperty("FirstOrderOfDataInDescription", dataTable.Rows[0]["FirstOrderOfDataInDescription"]),
                                                    new JProperty("SecondOrderOfDataInDescription", dataTable.Rows[0]["SecondOrderOfDataInDescription"]),
                                                    new JProperty("ThirdOrderOfDataInDescription", dataTable.Rows[0]["ThirdOrderOfDataInDescription"]),
                                                    new JProperty("FourthOrderOfDataInDescription", dataTable.Rows[0]["FourthOrderOfDataInDescription"]),
                                                    new JProperty("FifthOrderOfDataInDescription", dataTable.Rows[0]["FifthOrderOfDataInDescription"]));

                    if (!string.IsNullOrEmpty(dataTable.Rows[0]["AbbreviationFileName"].ToString()))
                    {
                        string AbbreviationFileName = dataTable.Rows[0]["AbbreviationFileName"].ToString();
                        string NewAbbreviationFileName = dataTable.Rows[0]["DGSettingID"].ToString() + " ~ " + AbbreviationFileName;
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GAT/Description Generator/"));
                        string tempFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + AbbreviationFileName);
                        if (File.Exists(dirUploads + NewAbbreviationFileName))
                        {
                            FileInfo fileAbbreviation = new FileInfo(dirUploads + NewAbbreviationFileName);
                            fileAbbreviation.CopyTo(tempFilepath);
                        }
                    }

                    if (!string.IsNullOrEmpty(dataTable.Rows[0]["IdentifierFileName"].ToString()))
                    {
                        string IdentifierFileName = dataTable.Rows[0]["IdentifierFileName"].ToString();
                        string NewIdentifierFileName = dataTable.Rows[0]["DGSettingID"].ToString() + " ~ " + IdentifierFileName;
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GAT/Description Generator/"));
                        string tempFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + IdentifierFileName);
                        if (File.Exists(dirUploads + NewIdentifierFileName))
                        {
                            FileInfo fileIdentifiers = new FileInfo(dirUploads + NewIdentifierFileName);
                            fileIdentifiers.CopyTo(tempFilepath);
                        }
                    }

                    return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }
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

        #region Helping Functions
        private string ReplaceAt(string str, int index, int length, string replace)
        {
            return str.Remove(index, Math.Min(length, str.Length - index))
                    .Insert(index, replace);
        }
        #endregion

        #region Get Abbreviated String
        private string GetAbbreviatedString(string InputString, List<clsAbbreviation> DGAbbreviationListData)
        {
            char NextCharacter, PreviousCharacter;
            string AbbreviatedString = string.Empty, ExpandedVersion = string.Empty, Abbreviation = string.Empty, AbbreviatedStringReplacedWithSpace = string.Empty;
            int indexOfExpandedVersion = -1;

            AbbreviatedString = InputString;
            foreach (var abb in DGAbbreviationListData.OrderByDescending(o => o.ExpandedVersion.Length))
            {
                #region Get the input string abbreviated
                if (AbbreviatedString.Contains(abb.ExpandedVersion))
                {
                    ExpandedVersion = abb.ExpandedVersion;
                    Abbreviation = abb.Abbreviation;

                    indexOfExpandedVersion = AbbreviatedString.IndexOf(ExpandedVersion);
                    while (true)
                    {
                        if (indexOfExpandedVersion < 0)
                            break;
                        else
                        {
                            if (indexOfExpandedVersion == 0)  //At Beginning
                            {
                                if (AbbreviatedString.ToUpper() == ExpandedVersion.ToUpper())
                                {
                                    AbbreviatedString = Abbreviation;
                                    break;
                                }

                                NextCharacter = AbbreviatedString[ExpandedVersion.Length];
                                if (!Char.IsLetter(NextCharacter))
                                    AbbreviatedString = ReplaceAt(AbbreviatedString, indexOfExpandedVersion, ExpandedVersion.Length, Abbreviation);
                                if (AbbreviatedString.Contains(ExpandedVersion))
                                    indexOfExpandedVersion = AbbreviatedString.IndexOf(ExpandedVersion, indexOfExpandedVersion + 1);
                                else
                                    break;
                            }
                            else if (indexOfExpandedVersion > 0 && AbbreviatedString.Length > indexOfExpandedVersion + ExpandedVersion.Length)  //Middle
                            {
                                NextCharacter = AbbreviatedString[indexOfExpandedVersion + ExpandedVersion.Length];
                                PreviousCharacter = AbbreviatedString[indexOfExpandedVersion - 1];
                                if (!Char.IsLetter(NextCharacter) && !Char.IsLetter(PreviousCharacter))
                                    AbbreviatedString = ReplaceAt(AbbreviatedString, indexOfExpandedVersion, ExpandedVersion.Length, Abbreviation);
                                if (AbbreviatedString.Contains(ExpandedVersion))
                                    indexOfExpandedVersion = AbbreviatedString.IndexOf(ExpandedVersion, indexOfExpandedVersion + 1);
                                else
                                    break;
                            }
                            else   //At End
                            {
                                PreviousCharacter = AbbreviatedString[indexOfExpandedVersion - 1];
                                if (Char.IsLetter(PreviousCharacter))
                                    break;
                                else
                                {
                                    AbbreviatedString = ReplaceAt(AbbreviatedString, indexOfExpandedVersion, ExpandedVersion.Length, Abbreviation);
                                    break;
                                }
                            }
                        }
                    }
                }
                #endregion
            }

            return AbbreviatedString;
        }
        #endregion

        #region Add MFR Details To Description
        private string[] AddMFRDetailsToDescription(Worksheet wsIF, int IFRow, string Description, string LeftOutInformation, int MFRNameColumnNo, bool IsMFRNameToBeAbbreviated, List<clsAbbreviation> DGAbbreviationListData, DescriptionGenerator model)
        {
            int MFRCounter = 1, CharacterLimit = 0;
            bool IsMFRNameExcluded = false, IsMFRPartNoExcluded = false;
            bool IncludeAllOtherMFRNames = false, IncludeAllOtherMFRPartNos = false, PrefixAllMFRNames = false, PrefixAllMFRPartNos = false;
            string MFRName = string.Empty, MFRPartNo = string.Empty;
            string PrefixMFRName = model.PrefixForMFRName;
            string PrefixMFRPartNo = model.PrefixForMFRPartNo;
            string DelimiterAfterMFRName = model.DelimiterAfterMFRName ?? ", ";
            string DelimiterAfterMFRPartNo = model.DelimiterAfterMFRPartNo ?? ", ";
            string[] returnArray = new string[2];

            if (model.IsMFRNameExcluded)
                IsMFRNameExcluded = true;

            if (model.IsMFRPartNoExcluded)
                IsMFRPartNoExcluded = true;

            if (model.IsToIncludeAllOtherMFRNames)
            {
                IncludeAllOtherMFRNames = true;
                if (model.IsToPrefixAllMFRNames)
                    PrefixAllMFRNames = true;
            }

            if (model.IsToIncludeAllOtherMFRPartNos)
            {
                IncludeAllOtherMFRPartNos = true;
                if (model.IsToPrefixAllMFRPartNos)
                    PrefixAllMFRPartNos = true;
            }

            if (model.DescriptionToGenerate == "S")
            {
                if (model.CharacterLimit > 0)
                    CharacterLimit = model.CharacterLimit;
            }

            while (MFRCounter <= 10)
            {
                MFRName = wsIF.Cells[IFRow, MFRNameColumnNo].StringValue;
                MFRPartNo = wsIF.Cells[IFRow, MFRNameColumnNo + 1].StringValue;
                if (!IsMFRNameExcluded)
                {
                    if (IsMFRNameToBeAbbreviated)
                    {
                        if (!string.IsNullOrEmpty(MFRName))
                            MFRName = GetAbbreviatedString(MFRName, DGAbbreviationListData);
                    }

                    if (!string.IsNullOrEmpty(MFRName))
                    {
                        if (MFRCounter > 1)
                        {
                            if (IncludeAllOtherMFRNames)
                            {
                                if (PrefixAllMFRNames)
                                {
                                    if (model.DescriptionToGenerate == "S")
                                    {
                                        if (Description.Length + PrefixMFRName.Length + MFRName.Length <= CharacterLimit)
                                            Description = Description + PrefixMFRName + MFRName + DelimiterAfterMFRName;
                                        else
                                            LeftOutInformation = LeftOutInformation + PrefixMFRName + MFRName + DelimiterAfterMFRName;
                                    }
                                    else
                                        Description = Description + PrefixMFRName + MFRName + DelimiterAfterMFRName;
                                }
                                else
                                {
                                    if (model.DescriptionToGenerate == "S")
                                    {
                                        if (Description.Length + MFRName.Length <= CharacterLimit)
                                            Description = Description + MFRName + DelimiterAfterMFRName;
                                        else
                                            LeftOutInformation = LeftOutInformation + MFRName + DelimiterAfterMFRName;
                                    }
                                    else
                                        Description = Description + MFRName + DelimiterAfterMFRName;
                                }
                            }
                        }
                        else
                        {
                            if (model.DescriptionToGenerate == "S")
                            {
                                if (Description.Length + PrefixMFRName.Length + MFRName.Length <= CharacterLimit)
                                    Description = Description + PrefixMFRName + MFRName + DelimiterAfterMFRName;
                                else
                                    LeftOutInformation = LeftOutInformation + PrefixMFRName + MFRName + DelimiterAfterMFRName;
                            }
                            else
                                Description = Description + PrefixMFRName + MFRName + DelimiterAfterMFRName;
                        }
                    }
                }

                if (!IsMFRPartNoExcluded)
                {
                    if (!string.IsNullOrEmpty(MFRPartNo))
                    {
                        if (MFRCounter > 1)
                        {
                            if (IncludeAllOtherMFRPartNos)
                            {
                                if (PrefixAllMFRPartNos)
                                {
                                    if (model.DescriptionToGenerate == "S")
                                    {
                                        if (Description.Length + PrefixMFRPartNo.Length + MFRPartNo.Length <= CharacterLimit)
                                            Description = Description + PrefixMFRPartNo + MFRPartNo + DelimiterAfterMFRPartNo;
                                        else
                                            LeftOutInformation = LeftOutInformation + PrefixMFRPartNo + MFRPartNo + DelimiterAfterMFRPartNo;
                                    }
                                    else
                                        Description = Description + PrefixMFRPartNo + MFRPartNo + DelimiterAfterMFRPartNo;
                                }
                                else
                                {
                                    if (model.DescriptionToGenerate == "S")
                                    {
                                        if (Description.Length + MFRPartNo.Length <= CharacterLimit)
                                            Description = Description + MFRPartNo + DelimiterAfterMFRPartNo;
                                        else
                                            LeftOutInformation = LeftOutInformation + MFRPartNo + DelimiterAfterMFRPartNo;
                                    }
                                    else
                                        Description = Description + MFRPartNo + DelimiterAfterMFRPartNo;
                                }
                            }
                        }
                        else
                        {
                            if (model.DescriptionToGenerate == "S")
                            {
                                if (Description.Length + PrefixMFRPartNo.Length + MFRPartNo.Length <= CharacterLimit)
                                    Description = Description + PrefixMFRPartNo + MFRPartNo + DelimiterAfterMFRPartNo;
                                else
                                    LeftOutInformation = LeftOutInformation + PrefixMFRPartNo + MFRPartNo + DelimiterAfterMFRPartNo;
                            }
                            else
                                Description = Description + PrefixMFRPartNo + MFRPartNo + DelimiterAfterMFRPartNo;
                        }
                    }
                }

                MFRCounter++;
                MFRNameColumnNo += 2;
                if (MFRCounter > 1)
                {
                    if (!IncludeAllOtherMFRNames && !IncludeAllOtherMFRPartNos)
                        break;
                }
            }

            returnArray[0] = Description;
            returnArray[1] = LeftOutInformation;

            return returnArray;
        }
        #endregion

        #region Applying Identifier to Attribute Value
        private string GetIdentifierAppliedAttributeValue(DescriptionGenerator model, string AttributeName, string AttributeValue, List<Identifier> IdentifierList)
        {
            bool IsToAddSpaceWhileApplyingIdentifier = false;

            if (model.IsToAddSpaceBeforeORAfterIdentifier)
                IsToAddSpaceWhileApplyingIdentifier = true;

            if (IdentifierList.Any(i => i.Characteristics.Trim().ToUpper() == AttributeName.Trim().ToUpper()))
            {
                Identifier identifier = IdentifierList.Where(i => i.Characteristics.Trim().ToUpper() == AttributeName.Trim().ToUpper()).First();
                if (identifier.SuffixORPrefix.Trim().ToUpper() == "SUFFIX")
                {
                    if (IsToAddSpaceWhileApplyingIdentifier)
                        AttributeValue = AttributeValue.Trim() + " " + identifier.Abbreviation.Trim();
                    else
                        AttributeValue = AttributeValue.Trim() + identifier.Abbreviation.Trim();
                }
                else
                {
                    if (IsToAddSpaceWhileApplyingIdentifier)
                        AttributeValue = identifier.Abbreviation.Trim() + " " + AttributeValue.Trim();
                    else
                        AttributeValue = identifier.Abbreviation.Trim() + AttributeValue.Trim();
                }
            }
            return AttributeValue;
        }
        #endregion

        #region Add Attribute Details To Description
        private string[] AddAttributeDetailsToDescription(Worksheet wsIF, int IFrow, string Description, string LeftOutInformation, int AttributeNameColumnNo, bool IsAttributeNameToBeAbbreviated, bool IsAttributeValueToBeAbbreviated, List<clsAbbreviation> DGAbbreviationListData, bool IsIdentifiersToBeApplied, List<Identifier> IdentifierList, DescriptionGenerator model)
        {
            int AttributeCounter = 1, CharacterLimit = 0;
            char MultipleAttributeValuesSeparator;
            bool IsAttributeNameExcluded = false, IsAttributeValueExcluded = false;
            string DelimiterAfterAttributeName = ": ", DelimiterAfterAttributeValue = ", ";
            bool IncludeAttributeNameWithNULLValues = false;
            bool IsToIncludeMaximumValues = false, IsToInterpretAllAttributeValues = false;
            string AttributeName = string.Empty, AttributeValue = string.Empty;
            string[] returnArray = new string[2];

            if (model.IsAttributeNameExcluded)
                IsAttributeNameExcluded = true;

            if (model.IsAttributeValueExcluded)
                IsAttributeValueExcluded = true;

            DelimiterAfterAttributeName = model.DelimiterAfterAttributeName ?? ": ";
            DelimiterAfterAttributeValue = model.DelimiterAfterAttributeValue ?? ", ";
            MultipleAttributeValuesSeparator = model.MultipleValuesSeparator != null ? Convert.ToChar(model.MultipleValuesSeparator) : ',';

            if (model.IsToIncludeAttributeNameWithNULLValues)
                IncludeAttributeNameWithNULLValues = true;

            if (model.IsToIncludeMaximumValues)
                IsToIncludeMaximumValues = true;

            if (model.IsToInterpretAllAttributeValues)
                IsToInterpretAllAttributeValues = true;

            if (model.DescriptionToGenerate == "S")
            {
                if (model.CharacterLimit > 0)
                    CharacterLimit = model.CharacterLimit;
            }

            bool IsAddedToLeftOutInformation = false;
            if (!(IsAttributeNameExcluded && IsAttributeValueExcluded))
            {
                while (AttributeCounter <= 50)
                {
                    AttributeName = wsIF.Cells[IFrow, AttributeNameColumnNo].StringValue.Trim();
                    AttributeValue = wsIF.Cells[IFrow, AttributeNameColumnNo + 1].StringValue.Trim();

                    if (string.IsNullOrEmpty(AttributeName) && string.IsNullOrEmpty(AttributeValue))
                        break;

                    #region Abbreviate and Apply Identifiers To Attribute Value
                    if (!IsAttributeValueExcluded)
                    {
                        if (!string.IsNullOrEmpty(AttributeValue))
                        {
                            if (IsAttributeValueToBeAbbreviated)
                                AttributeValue = GetAbbreviatedString(AttributeValue, DGAbbreviationListData);

                            if (IsIdentifiersToBeApplied)
                            {
                                if (!string.IsNullOrEmpty(AttributeName) && !string.IsNullOrEmpty(AttributeValue))
                                    AttributeValue = GetIdentifierAppliedAttributeValue(model, AttributeName, AttributeValue, IdentifierList);
                            }
                        }
                    }
                    #endregion

                    #region Abbreviate Attribute Name
                    if (!IsAttributeNameExcluded)
                    {
                        if (!string.IsNullOrEmpty(AttributeName))
                        {
                            if (IsAttributeNameToBeAbbreviated)
                                AttributeName = GetAbbreviatedString(AttributeName, DGAbbreviationListData);
                        }
                    }
                    #endregion

                    #region Generate Description
                    if (model.DescriptionToGenerate == "S")
                    {
                        #region Generate Short Description
                        if (!(string.IsNullOrEmpty(AttributeName) && string.IsNullOrEmpty(AttributeValue)))
                        {
                            if (!IsAttributeNameExcluded && IsAttributeValueExcluded)
                            {
                                #region Add only Attribute Name
                                if (!string.IsNullOrEmpty(AttributeName))
                                {
                                    if (AttributeName.Length + Description.Length <= CharacterLimit)
                                    {
                                        if (IncludeAttributeNameWithNULLValues)
                                            Description = Description + AttributeName + DelimiterAfterAttributeName;
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(AttributeValue))
                                                Description = Description + AttributeName + DelimiterAfterAttributeName;
                                        }
                                    }
                                    else
                                    {
                                        if (IncludeAttributeNameWithNULLValues)
                                            LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName;
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(AttributeValue))
                                                LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName;
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (IsAttributeNameExcluded && !IsAttributeValueExcluded)
                            {
                                #region Add only Attribute Value
                                if (!string.IsNullOrEmpty(AttributeValue))
                                {
                                    if (IsToInterpretAllAttributeValues)
                                    {
                                        string[] AttributeValueArray = AttributeValue.Split(MultipleAttributeValuesSeparator);
                                        foreach (string tav in AttributeValueArray)
                                        {
                                            if (IsToIncludeMaximumValues)
                                            {
                                                if (tav.Trim().Length + Description.Length <= CharacterLimit)
                                                    Description = Description + tav.Trim() + MultipleAttributeValuesSeparator;
                                                else
                                                    LeftOutInformation = LeftOutInformation + tav.Trim() + MultipleAttributeValuesSeparator;
                                            }
                                            else
                                            {
                                                if (!IsAddedToLeftOutInformation)
                                                {
                                                    if (tav.Trim().Length + Description.Length <= CharacterLimit)
                                                        Description = Description + tav.Trim() + MultipleAttributeValuesSeparator;
                                                    else
                                                    {
                                                        LeftOutInformation = LeftOutInformation + tav.Trim() + MultipleAttributeValuesSeparator;
                                                        IsAddedToLeftOutInformation = true;
                                                    }
                                                }
                                                else
                                                    LeftOutInformation = LeftOutInformation + tav.Trim() + MultipleAttributeValuesSeparator;
                                            }
                                        }

                                        if (!string.IsNullOrEmpty(Description))
                                        {
                                            if (Description.Substring(Description.Length - 1, 1) == MultipleAttributeValuesSeparator.ToString())
                                                Description = ReplaceAt(Description, Description.Length - 1, 1, DelimiterAfterAttributeValue);
                                        }

                                        if (!string.IsNullOrEmpty(LeftOutInformation))
                                        {
                                            if (LeftOutInformation.Substring(LeftOutInformation.Length - 1, 1) == MultipleAttributeValuesSeparator.ToString())
                                                LeftOutInformation = ReplaceAt(LeftOutInformation, LeftOutInformation.Length - 1, 1, DelimiterAfterAttributeValue);
                                        }
                                    }
                                    else
                                    {
                                        if (IsToIncludeMaximumValues)
                                        {
                                            if (AttributeValue.Length + Description.Length <= CharacterLimit)
                                                Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                                            else
                                                LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                        }
                                        else
                                        {
                                            if (!IsAddedToLeftOutInformation)
                                            {
                                                if (AttributeValue.Length + Description.Length <= CharacterLimit)
                                                    Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                                                else
                                                {
                                                    LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                    IsAddedToLeftOutInformation = true;
                                                }
                                            }
                                            else
                                                LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                        }
                                    }
                                }
                                #endregion
                            }
                            else if (!IsAttributeNameExcluded && !IsAttributeValueExcluded)
                            {
                                #region Add both Attribute Name and Attribute Value
                                if (!IsToIncludeMaximumValues && IsAddedToLeftOutInformation)
                                {
                                    #region flag is OFF and Is Added To LeftOutInformation
                                    if (!string.IsNullOrEmpty(AttributeName))
                                    {
                                        if (IncludeAttributeNameWithNULLValues)
                                            LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(AttributeValue))
                                                LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region Either include max. value is ON or Left out is empty
                                    if (AttributeName.Length + DelimiterAfterAttributeName.Length + AttributeValue.Length + Description.Length <= CharacterLimit)
                                    {
                                        if (!string.IsNullOrEmpty(AttributeName))
                                        {
                                            if (IncludeAttributeNameWithNULLValues)
                                                Description = Description + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                            else
                                            {
                                                if (!string.IsNullOrEmpty(AttributeValue))
                                                    Description = Description + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (IsToInterpretAllAttributeValues)
                                        {
                                            if (!string.IsNullOrEmpty(AttributeValue))
                                            {
                                                string[] AttributeValueArray = AttributeValue.Split(MultipleAttributeValuesSeparator);
                                                string tempDesc = string.Empty, tempLeftOutInfo = string.Empty;
                                                foreach (string tav in AttributeValueArray)
                                                {
                                                    if (AttributeName.Length + DelimiterAfterAttributeName.Length + tav.Trim().Length + Description.Length <= CharacterLimit)
                                                        tempDesc = tempDesc + tav.Trim() + MultipleAttributeValuesSeparator;
                                                    else
                                                        tempLeftOutInfo = tempLeftOutInfo + tav.Trim() + MultipleAttributeValuesSeparator;
                                                }

                                                if (!string.IsNullOrEmpty(tempDesc))
                                                    Description = Description + AttributeName + DelimiterAfterAttributeName + tempDesc;

                                                if (!string.IsNullOrEmpty(tempLeftOutInfo))
                                                {
                                                    LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + tempLeftOutInfo;
                                                    IsAddedToLeftOutInformation = true;
                                                }

                                                if (!string.IsNullOrEmpty(Description))
                                                {
                                                    if (Description.Substring(Description.Length - 1, 1) == MultipleAttributeValuesSeparator.ToString())
                                                        Description = ReplaceAt(Description, Description.Length - 1, 1, DelimiterAfterAttributeValue);
                                                }

                                                if (!string.IsNullOrEmpty(LeftOutInformation))
                                                {
                                                    if (LeftOutInformation.Substring(LeftOutInformation.Length - 1, 1) == MultipleAttributeValuesSeparator.ToString())
                                                        LeftOutInformation = ReplaceAt(LeftOutInformation, LeftOutInformation.Length - 1, 1, DelimiterAfterAttributeValue);
                                                }
                                            }
                                            else
                                            {
                                                if (IncludeAttributeNameWithNULLValues)
                                                {
                                                    if (AttributeName.Length + DelimiterAfterAttributeName.Length + Description.Length <= CharacterLimit)
                                                        Description = Description + AttributeName + DelimiterAfterAttributeName;
                                                    else
                                                    {
                                                        LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName;
                                                        IsAddedToLeftOutInformation = true;
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (!string.IsNullOrEmpty(AttributeName))
                                            {
                                                if (IncludeAttributeNameWithNULLValues)
                                                {
                                                    LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                    IsAddedToLeftOutInformation = true;
                                                }
                                                else
                                                {
                                                    if (!string.IsNullOrEmpty(AttributeValue))
                                                    {
                                                        LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                        IsAddedToLeftOutInformation = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region Generate Long Description
                        if (!IsAttributeNameExcluded)
                        {
                            if (!string.IsNullOrEmpty(AttributeName))
                            {
                                if (IncludeAttributeNameWithNULLValues)
                                    Description = Description + AttributeName + DelimiterAfterAttributeName;
                                else
                                {
                                    if (!string.IsNullOrEmpty(AttributeValue))
                                        Description = Description + AttributeName + DelimiterAfterAttributeName;
                                }
                            }
                        }

                        if (!IsAttributeValueExcluded)
                        {
                            if (!string.IsNullOrEmpty(AttributeValue))
                                Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                        }
                        #endregion
                    }
                    #endregion

                    AttributeCounter++;
                    AttributeNameColumnNo += 2;
                }
                returnArray[0] = Description;
                returnArray[1] = LeftOutInformation;
            }
            return returnArray;
        }
        #endregion

        #region Add Additional Information To Description
        private string[] AddAdditionalInformationToDescription(Worksheet wsIF, DescriptionGenerator model, int IFrow, string Description, string LeftOutInformation, int AdditionalInfoColumnNo, bool IsAttributeNameToBeAbbreviated, bool IsAttributeValueToBeAbbreviated, bool IsAdditionalInformationToBeAbbreviated, List<clsAbbreviation> AbbreviationList, bool IsIdentifiersToBeApplied, List<Identifier> IdentifierList)
        {
            bool IsAdditionalInformationToBeInterpreted = false, IsToIncludeAttributeNameFromAdditionalInformation = false;
            int CharacterLimit = 0;
            bool IsAttributeValueExcluded = false, IsAdditionalInfoExcluded = false;
            string DelimiterAfterAttributeName = ": ", DelimiterAfterAttributeValue = ", ", DelimiterAfterAdditionalInformation = ", ";
            bool IncludeAttributeNameWithNULLValues = false;
            bool IsToIncludeMaximumValues = false;
            string AttributeName = string.Empty, AttributeValue = string.Empty, AdditionalInformation = string.Empty;
            bool IsIdentifiersToBeAppliedToAdditonalInformation = false;
            string PrefixAdditionalInformation = model.PrefixForAdditionalInformation;
            string[] returnArray = new string[2];

            if (model.IsAdditionalInformationExcluded)
                IsAdditionalInfoExcluded = true;

            if (!IsAdditionalInfoExcluded)
            {
                #region Assigning the control values
                if (model.IsAttributeValueExcluded)
                    IsAttributeValueExcluded = true;

                if (model.IsToInterpretAdditionalInformation)
                    IsAdditionalInformationToBeInterpreted = true;

                if (model.IsToIncludeAttributeNameFromAdditionalInformation)
                    IsToIncludeAttributeNameFromAdditionalInformation = true;

                DelimiterAfterAttributeName = model.DelimiterAfterAttributeName ?? ": ";
                DelimiterAfterAttributeValue = model.DelimiterAfterAttributeValue ?? ", ";
                DelimiterAfterAdditionalInformation = model.DelimiterAfterAdditionalInformation ?? ", ";

                if (model.IsToIncludeAttributeNameWithNULLValues)
                    IncludeAttributeNameWithNULLValues = true;

                if (model.IsToIncludeMaximumValues)
                    IsToIncludeMaximumValues = true;

                if (model.IsToApplyIdentifierToAdditionalInformation)
                    IsIdentifiersToBeAppliedToAdditonalInformation = true;

                if (model.DescriptionToGenerate == "S")
                {
                    if (model.CharacterLimit > 0)
                        CharacterLimit = model.CharacterLimit;
                }
                #endregion

                AdditionalInformation = wsIF.Cells[IFrow, AdditionalInfoColumnNo].StringValue;
                if (!string.IsNullOrEmpty(AdditionalInformation))
                {
                    if (IsAdditionalInformationToBeInterpreted)
                    {
                        #region Interpret Additional Information, Abbreviate it, Apply Identifier, and generate Short or Long description
                        string AddInfo = AdditionalInformation, AddInfotmp = AdditionalInformation;
                        char charAtIndex;
                        string AttributeNameDelimiter = DelimiterAfterAttributeName, AttributeValueDelimiter = DelimiterAfterAttributeValue, MultipleValuesSeparator = model.MultipleValuesSeparator;
                        string charsAddedString = string.Empty;
                        int AttributeValueCounter = 0;
                        bool IsToFireTheLogic = false, IsPrefixAdded = false, IsAddedToLeftOutInformation = false;
                        string PreviousAttributeName = string.Empty, MultipleValuesTempString = string.Empty;

                        for (int i = 0; i < AddInfotmp.Length; i++)
                        {
                            if (AddInfo.Length > 0)
                            {
                                #region Extracting Attribute Name and Value from Additional Information string
                                charAtIndex = AddInfo[0];
                                if (charAtIndex.ToString() != AttributeNameDelimiter.Trim() && charAtIndex.ToString() != AttributeValueDelimiter.Trim() && charAtIndex.ToString() != MultipleValuesSeparator.Trim())
                                {
                                    charsAddedString = charsAddedString + charAtIndex.ToString();
                                    AddInfo = ReplaceAt(AddInfo, 0, 1, string.Empty);
                                    IsToFireTheLogic = false;
                                }

                                if (charAtIndex.ToString() == AttributeNameDelimiter.Trim())
                                {
                                    AttributeName = charsAddedString;
                                    charsAddedString = string.Empty;
                                    AddInfo = ReplaceAt(AddInfo, 0, AttributeNameDelimiter.Length, string.Empty);
                                    AttributeValueCounter = 0;
                                    IsToFireTheLogic = false;
                                }

                                if (charAtIndex.ToString() == MultipleValuesSeparator.Trim())
                                {
                                    AttributeValue = charsAddedString;
                                    charsAddedString = string.Empty;
                                    AddInfo = ReplaceAt(AddInfo, 0, MultipleValuesSeparator.Length, string.Empty);
                                    IsToFireTheLogic = true;
                                }

                                if (charAtIndex.ToString() == AttributeValueDelimiter.Trim() || AddInfo.Length == 0)
                                {
                                    AttributeValue = charsAddedString;
                                    charsAddedString = string.Empty;
                                    AddInfo = ReplaceAt(AddInfo, 0, AttributeValueDelimiter.Length, string.Empty);
                                    AttributeValueCounter++;
                                    if (AttributeValueCounter > 1)
                                        AttributeName = string.Empty;

                                    IsToFireTheLogic = true;
                                }
                                #endregion

                                if (IsToFireTheLogic)
                                {
                                    if (model.DescriptionToGenerate == "S")
                                    {
                                        #region Generate Short Description
                                        AttributeName = AttributeName.Trim();
                                        AttributeValue = AttributeValue.Trim();
                                        if (IsToIncludeAttributeNameFromAdditionalInformation && IsAttributeValueExcluded)
                                        {
                                            #region Abbreviate Attribute Name
                                            if (!string.IsNullOrEmpty(AttributeName))
                                            {
                                                if (IsAttributeNameToBeAbbreviated)
                                                    AttributeName = GetAbbreviatedString(AttributeName, AbbreviationList);
                                            }
                                            #endregion

                                            #region Add only Attribute Name
                                            if (!string.IsNullOrEmpty(AttributeName))
                                            {
                                                if (IsPrefixAdded)
                                                {
                                                    if (AttributeName.Length + Description.Length <= CharacterLimit)
                                                        Description = Description + AttributeName + DelimiterAfterAttributeName;
                                                    else
                                                        LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName;
                                                }
                                                else
                                                {
                                                    if (AttributeName.Length + PrefixAdditionalInformation.Length + Description.Length <= CharacterLimit)
                                                    {
                                                        Description = Description + PrefixAdditionalInformation + AttributeName + DelimiterAfterAttributeName;
                                                        IsPrefixAdded = true;
                                                    }
                                                    else
                                                        LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName;
                                                }
                                            }
                                            #endregion
                                        }
                                        else if (!IsToIncludeAttributeNameFromAdditionalInformation && !IsAttributeValueExcluded)
                                        {
                                            #region Apply Identifier
                                            if (IsIdentifiersToBeAppliedToAdditonalInformation)
                                            {
                                                if (!string.IsNullOrEmpty(AttributeName) && !string.IsNullOrEmpty(AttributeValue) && charAtIndex.ToString().ToUpper() == DelimiterAfterAttributeValue.Trim().ToUpper())
                                                    AttributeValue = GetIdentifierAppliedAttributeValue(model, AttributeName, AttributeValue, IdentifierList);
                                            }
                                            #endregion

                                            #region Abbreviate Attribute Value
                                            if (!string.IsNullOrEmpty(AttributeValue))
                                            {
                                                if (IsAttributeValueToBeAbbreviated)
                                                    AttributeValue = GetAbbreviatedString(AttributeValue, AbbreviationList);
                                            }
                                            #endregion

                                            #region Add only Attribute Value
                                            if (!string.IsNullOrEmpty(AttributeValue))
                                            {
                                                //ACCURACY CLASS: 1.5,2.4,4.5; Opposite of BRAND: METALITE HANDY,J155
                                                if (AddInfo.Length == 0 || charAtIndex.ToString() == DelimiterAfterAttributeValue.Trim())
                                                {
                                                    if (!string.IsNullOrEmpty(MultipleValuesTempString))
                                                    {
                                                        AttributeValue = MultipleValuesTempString + AttributeValue;
                                                        MultipleValuesTempString = string.Empty;
                                                    }

                                                    #region Process Attribute Value
                                                    if (IsToIncludeMaximumValues)
                                                    {
                                                        if (IsPrefixAdded)
                                                        {
                                                            if (AttributeValue.Length + Description.Length <= CharacterLimit)
                                                                Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                                                            else
                                                                LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                        }
                                                        else
                                                        {
                                                            if (AttributeValue.Length + PrefixAdditionalInformation.Length + Description.Length <= CharacterLimit)
                                                            {
                                                                Description = Description + PrefixAdditionalInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                                IsPrefixAdded = true;
                                                            }
                                                            else
                                                                LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (IsAddedToLeftOutInformation)
                                                            LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                        else
                                                        {
                                                            if (IsPrefixAdded)
                                                            {
                                                                if (AttributeValue.Length + Description.Length <= CharacterLimit)
                                                                    Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                                                                else
                                                                {
                                                                    LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                                    IsAddedToLeftOutInformation = true;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (AttributeValue.Length + PrefixAdditionalInformation.Length + Description.Length <= CharacterLimit)
                                                                {
                                                                    Description = Description + PrefixAdditionalInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                                    IsPrefixAdded = true;
                                                                }
                                                                else
                                                                {
                                                                    LeftOutInformation = LeftOutInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                                    IsAddedToLeftOutInformation = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    MultipleValuesTempString = MultipleValuesTempString + AttributeValue + MultipleValuesSeparator;
                                                }
                                            }
                                            #endregion
                                        }
                                        else if (IsToIncludeAttributeNameFromAdditionalInformation && !IsAttributeValueExcluded)
                                        {
                                            #region Apply Identifier
                                            if (IsIdentifiersToBeAppliedToAdditonalInformation)
                                            {
                                                if (!string.IsNullOrEmpty(AttributeName) && !string.IsNullOrEmpty(AttributeValue) && charAtIndex.ToString().ToUpper() == DelimiterAfterAttributeValue.Trim().ToUpper())
                                                    AttributeValue = GetIdentifierAppliedAttributeValue(model, AttributeName, AttributeValue, IdentifierList);
                                            }
                                            #endregion

                                            #region Abbreviate Attribute Name
                                            if (!string.IsNullOrEmpty(AttributeName))
                                            {
                                                if (IsAttributeNameToBeAbbreviated)
                                                    AttributeName = GetAbbreviatedString(AttributeName, AbbreviationList);
                                            }
                                            #endregion

                                            #region Abbreviate Attribute Value
                                            if (!string.IsNullOrEmpty(AttributeValue))
                                            {
                                                if (IsAttributeValueToBeAbbreviated)
                                                    AttributeValue = GetAbbreviatedString(AttributeValue, AbbreviationList);
                                            }
                                            #endregion

                                            if (!(!IncludeAttributeNameWithNULLValues && string.IsNullOrEmpty(AttributeValue))) //Attribute Name Having Values
                                            {
                                                #region Add Attribute Name and Attribute Value both
                                                if (AddInfo.Length == 0 || charAtIndex.ToString() == DelimiterAfterAttributeValue.Trim())
                                                {
                                                    if (!string.IsNullOrEmpty(MultipleValuesTempString))
                                                    {
                                                        AttributeValue = MultipleValuesTempString + AttributeValue;
                                                        MultipleValuesTempString = string.Empty;
                                                    }

                                                    #region Different Attributes or Values
                                                    if (IsToIncludeMaximumValues)
                                                    {
                                                        if (IsPrefixAdded)
                                                        {
                                                            if (AttributeName.Length + DelimiterAfterAttributeName.Length + AttributeValue.Length + Description.Length <= CharacterLimit)
                                                                Description = Description + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                            else
                                                                LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                        }
                                                        else
                                                        {
                                                            if (AttributeName.Length + DelimiterAfterAttributeName.Length + AttributeValue.Length + PrefixAdditionalInformation.Length + Description.Length <= CharacterLimit)
                                                            {
                                                                Description = Description + PrefixAdditionalInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                                IsPrefixAdded = true;
                                                            }
                                                            else
                                                                LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (IsPrefixAdded)
                                                        {
                                                            if (IsAddedToLeftOutInformation)
                                                                LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                            else
                                                            {
                                                                if (AttributeName.Length + DelimiterAfterAttributeName.Length + AttributeValue.Length + Description.Length <= CharacterLimit)
                                                                    Description = Description + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                                else
                                                                {
                                                                    LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                                    IsAddedToLeftOutInformation = true;
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (IsAddedToLeftOutInformation)
                                                                LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                            else
                                                            {
                                                                if (AttributeName.Length + DelimiterAfterAttributeName.Length + AttributeValue.Length + PrefixAdditionalInformation.Length + Description.Length <= CharacterLimit)
                                                                {
                                                                    Description = Description + PrefixAdditionalInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                                    IsPrefixAdded = true;
                                                                }
                                                                else
                                                                {
                                                                    LeftOutInformation = LeftOutInformation + AttributeName + DelimiterAfterAttributeName + AttributeValue + DelimiterAfterAttributeValue;
                                                                    IsAddedToLeftOutInformation = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion
                                                }
                                                else
                                                {
                                                    MultipleValuesTempString = MultipleValuesTempString + AttributeValue + MultipleValuesSeparator;
                                                }
                                                #endregion
                                            }
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        #region Generate Long Description
                                        AttributeName = AttributeName.Trim();
                                        AttributeValue = AttributeValue.Trim();
                                        #region Apply Identifier
                                        if (IsIdentifiersToBeAppliedToAdditonalInformation)
                                        {
                                            if (!string.IsNullOrEmpty(AttributeName) && !string.IsNullOrEmpty(AttributeValue) && charAtIndex.ToString().ToUpper() == DelimiterAfterAttributeValue.Trim().ToUpper())
                                                AttributeValue = GetIdentifierAppliedAttributeValue(model, AttributeName, AttributeValue, IdentifierList);
                                        }
                                        #endregion

                                        #region Abbreviate Attribute Name
                                        if (!string.IsNullOrEmpty(AttributeName))
                                        {
                                            if (IsAttributeNameToBeAbbreviated)
                                                AttributeName = GetAbbreviatedString(AttributeName, AbbreviationList);
                                        }
                                        #endregion

                                        #region Abbreviate Attribute Value
                                        if (!string.IsNullOrEmpty(AttributeValue))
                                        {
                                            if (IsAttributeValueToBeAbbreviated)
                                                AttributeValue = GetAbbreviatedString(AttributeValue, AbbreviationList);
                                        }
                                        #endregion

                                        if (IsPrefixAdded)
                                        {
                                            if (PreviousAttributeName != AttributeName)
                                            {
                                                if (IsToIncludeAttributeNameFromAdditionalInformation)
                                                {
                                                    if (!string.IsNullOrEmpty(AttributeName))
                                                    {
                                                        if (IncludeAttributeNameWithNULLValues)
                                                            Description = Description + AttributeName + DelimiterAfterAttributeName;
                                                        else
                                                        {
                                                            if (!string.IsNullOrEmpty(AttributeValue))
                                                                Description = Description + AttributeName + DelimiterAfterAttributeName;
                                                        }
                                                    }
                                                }

                                                if (!IsAttributeValueExcluded)
                                                {
                                                    if (!string.IsNullOrEmpty(AttributeValue))
                                                    {
                                                        if (charAtIndex.ToString() == DelimiterAfterAttributeValue.Trim() || AddInfo.Length == 0)
                                                            Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                                                        else
                                                            Description = Description + AttributeValue + MultipleValuesSeparator;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (charAtIndex.ToString() == DelimiterAfterAttributeValue.Trim() || AddInfo.Length == 0)
                                                {
                                                    if (!IsAttributeValueExcluded)
                                                    {
                                                        if (!string.IsNullOrEmpty(AttributeValue))
                                                            Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                                                    }
                                                }
                                                else
                                                {
                                                    if (!IsAttributeValueExcluded)
                                                    {
                                                        if (!string.IsNullOrEmpty(AttributeValue))
                                                            Description = Description + AttributeValue + MultipleValuesSeparator;
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (IsToIncludeAttributeNameFromAdditionalInformation)
                                            {
                                                if (!string.IsNullOrEmpty(AttributeName))
                                                {
                                                    if (IncludeAttributeNameWithNULLValues)
                                                    {
                                                        Description = Description + PrefixAdditionalInformation + AttributeName + DelimiterAfterAttributeName;
                                                        IsPrefixAdded = true;
                                                    }
                                                    else
                                                    {
                                                        if (!string.IsNullOrEmpty(AttributeValue))
                                                        {
                                                            Description = Description + PrefixAdditionalInformation + AttributeName + DelimiterAfterAttributeName;
                                                            IsPrefixAdded = true;
                                                        }
                                                    }
                                                }
                                            }

                                            if (!IsAttributeValueExcluded)
                                            {
                                                if (!string.IsNullOrEmpty(AttributeValue))
                                                {
                                                    if (IsPrefixAdded)
                                                    {
                                                        if (charAtIndex.ToString() == DelimiterAfterAttributeValue.Trim() || AddInfo.Length == 0)
                                                            Description = Description + AttributeValue + DelimiterAfterAttributeValue;
                                                        else
                                                            Description = Description + AttributeValue + MultipleValuesSeparator;
                                                    }
                                                    else
                                                    {
                                                        if (charAtIndex.ToString() == DelimiterAfterAttributeValue.Trim() || AddInfo.Length == 0)
                                                        {
                                                            Description = Description + PrefixAdditionalInformation + AttributeValue + DelimiterAfterAttributeValue;
                                                            IsPrefixAdded = true;
                                                        }
                                                        else
                                                        {
                                                            Description = Description + PrefixAdditionalInformation + AttributeValue + MultipleValuesSeparator;
                                                            IsPrefixAdded = true;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion
                                        PreviousAttributeName = AttributeName;
                                    }
                                }
                            }
                        }

                        if (model.DescriptionToGenerate == "S")
                        {
                            if (IsToIncludeAttributeNameFromAdditionalInformation && IsAttributeValueExcluded)
                            {
                                if (!IsPrefixAdded)
                                    LeftOutInformation = PrefixAdditionalInformation + LeftOutInformation;
                                else
                                    Description = Description + DelimiterAfterAdditionalInformation;

                                if (LeftOutInformation.Length > 0 && LeftOutInformation.LastIndexOf(DelimiterAfterAttributeName) > -1)
                                    LeftOutInformation = LeftOutInformation + DelimiterAfterAdditionalInformation;
                            }
                            else if ((!IsToIncludeAttributeNameFromAdditionalInformation && !IsAttributeValueExcluded) || (IsToIncludeAttributeNameFromAdditionalInformation && !IsAttributeValueExcluded))
                            {
                                if (!IsPrefixAdded)
                                    LeftOutInformation = PrefixAdditionalInformation + LeftOutInformation;

                                if (Description.Length > 0 && Description.IndexOf(DelimiterAfterAttributeValue) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterAttributeValue), DelimiterAfterAttributeValue.Length).ToUpper() == DelimiterAfterAttributeValue.ToUpper() &&
                                        Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAttributeValue)).Length + DelimiterAfterAttributeValue.Length)
                                    {
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterAttributeValue.Length, DelimiterAfterAttributeValue.Length, DelimiterAfterAdditionalInformation);
                                    }
                                }

                                if (LeftOutInformation.Length > 0 && LeftOutInformation.IndexOf(DelimiterAfterAttributeValue) > -1)
                                {
                                    if (LeftOutInformation.Substring(LeftOutInformation.LastIndexOf(DelimiterAfterAttributeValue), DelimiterAfterAttributeValue.Length).ToUpper() == DelimiterAfterAttributeValue.ToUpper() &&
                                        LeftOutInformation.Length == LeftOutInformation.Substring(0, LeftOutInformation.LastIndexOf(DelimiterAfterAttributeValue)).Length + DelimiterAfterAttributeValue.Length)
                                    {
                                        LeftOutInformation = ReplaceAt(LeftOutInformation, LeftOutInformation.Length - DelimiterAfterAttributeValue.Length, DelimiterAfterAttributeValue.Length, DelimiterAfterAdditionalInformation);
                                    }
                                }
                            }
                        }
                        else
                        {
                            #region Replacing the delimiter at the end (either Attribute Name or Attribute Value) with additional information delimiter
                            if (Description.Length > 0 && Description.LastIndexOf(DelimiterAfterAttributeName) > -1)
                            {
                                if (Description.Substring(Description.LastIndexOf(DelimiterAfterAttributeName), DelimiterAfterAttributeName.Length).ToUpper() == DelimiterAfterAttributeName.ToUpper() &&
                                    Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAttributeName)).Length + DelimiterAfterAttributeName.Length)
                                {
                                    Description = ReplaceAt(Description, Description.Length - DelimiterAfterAttributeName.Length, DelimiterAfterAttributeName.Length, DelimiterAfterAdditionalInformation);
                                }
                            }

                            if (Description.Length > 0 && Description.LastIndexOf(DelimiterAfterAttributeValue) > -1)
                            {
                                if (Description.Substring(Description.LastIndexOf(DelimiterAfterAttributeValue), DelimiterAfterAttributeValue.Length).ToUpper() == DelimiterAfterAttributeValue.ToUpper() &&
                                    Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAttributeValue)).Length + DelimiterAfterAttributeValue.Length)
                                {
                                    Description = ReplaceAt(Description, Description.Length - DelimiterAfterAttributeValue.Length, DelimiterAfterAttributeValue.Length, DelimiterAfterAdditionalInformation);
                                }
                            }
                            #endregion
                        }
                        #endregion
                    }
                    else
                    {
                        #region Do not Interpret Additional Information, Abbreviate it and generate Short or Long Description
                        if (IsAdditionalInformationToBeAbbreviated)
                            AdditionalInformation = GetAbbreviatedString(AdditionalInformation, AbbreviationList);

                        if (model.DescriptionToGenerate == "S")
                        {
                            if (PrefixAdditionalInformation.Length + AdditionalInformation.Length + Description.Length <= CharacterLimit)
                                Description = Description + PrefixAdditionalInformation + AdditionalInformation + DelimiterAfterAdditionalInformation;
                            else
                                LeftOutInformation = LeftOutInformation + PrefixAdditionalInformation + AdditionalInformation + DelimiterAfterAdditionalInformation;
                        }
                        else
                            Description = Description + PrefixAdditionalInformation + AdditionalInformation + DelimiterAfterAdditionalInformation;
                        #endregion
                    }
                }
            }

            returnArray[0] = Description;
            returnArray[1] = LeftOutInformation;
            return returnArray;
        }
        #endregion

        #region Generate Description
        [HttpPost]
        [Route("GenerateDescription")]
        public HttpResponseMessage GenerateDescription([FromBody] DescriptionGenerator model)
        {
            bool IsNounExcluded = false, IsModifierExcluded = false, IsMFRNameExcluded = false, IsMFRPartNoExcluded = false;
            bool IsAttributeNameExcluded = false, IsAttributeValueExcluded = false, IsAdditionalInformationExcluded = false;
            bool IsNounToBeAbbreviated = false, IsModifierToBeAbbreviated = false, IsAttributeNameToBeAbbreviated = false;
            bool IsAttributeValueToBeAbbreviated = false, IsAdditionalInformationToBeAbbreviated = false, IsMFRNameToBeAbbreviated = false;
            string DelimiterAfterNoun = ", ", DelimiterAfterModifier = ", ";
            string DelimiterAfterMFRName = model.DelimiterAfterMFRName != null ? model.DelimiterAfterMFRName : ", ";
            string DelimiterAfterMFRPartNo = model.DelimiterAfterMFRPartNo != null ? model.DelimiterAfterMFRPartNo : ", ";
            string DelimiterAfterAttributeName = model.DelimiterAfterAttributeName != null ? model.DelimiterAfterAttributeName : ": ";
            string DelimiterAfterAttributeValue = model.DelimiterAfterAttributeValue != null ? model.DelimiterAfterAttributeValue : ", ";
            string DelimiterAfterAdditionalInformation = model.DelimiterAfterAdditionalInformation != null ? model.DelimiterAfterAdditionalInformation : ", ";
            char MultipleAttributeValuesSeparator = model.MultipleValuesSeparator != null ? Convert.ToChar(model.MultipleValuesSeparator) : ',';
            char TruncationType = 'M';
            int CharacterLimit = 0;
            char DelimiterForTruncation = ',';
            string ThirdOrderOfDataInDescription = string.Empty, FourthOrderOfDataInDescription = string.Empty, FifthOrderOfDataInDescription = string.Empty;
            bool IsIdentifiersToBeApplied = false;

            int NounColumnNo = -1, ModifierColumnNo = -1, MFRName1ColumnNo = -1, AttributeName1ColumnNo = -1, AdditionalInfoColumnNo = -1, DescriptionColumnNo = -1;
            string Noun = string.Empty, Modifier = string.Empty;
            string Description = string.Empty, LeftOutInformation = string.Empty;

            string AttributeFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + model.InputFileName), AbbreviationFilePath = string.Empty, IdentifierFilePath = string.Empty;

            try
            {
                #region Assigning the control values
                if (model.IsNounExcluded)
                    IsNounExcluded = true;

                if (model.IsModifierExcluded)
                    IsModifierExcluded = true;

                if (model.IsMFRNameExcluded)
                    IsMFRNameExcluded = true;

                if (model.IsMFRPartNoExcluded)
                    IsMFRPartNoExcluded = true;

                if (model.IsAttributeNameExcluded)
                    IsAttributeNameExcluded = true;

                if (model.IsAttributeValueExcluded)
                    IsAttributeValueExcluded = true;

                if (model.IsAdditionalInformationExcluded)
                    IsAdditionalInformationExcluded = true;

                if (model.IsNounToBeAbbreviated)
                    IsNounToBeAbbreviated = true;

                if (model.IsModifierToBeAbbreviated)
                    IsModifierToBeAbbreviated = true;

                if (model.IsAttributeNameToBeAbbreviated)
                    IsAttributeNameToBeAbbreviated = true;

                if (model.IsAttributeValueToBeAbbreviated)
                    IsAttributeValueToBeAbbreviated = true;

                if (model.IsAdditionalInformationToBeAbbreviated)
                    IsAdditionalInformationToBeAbbreviated = true;

                if (model.IsMFRNameToBeAbbreviated)
                    IsMFRNameToBeAbbreviated = true;

                if (model.IsToApplyIdentifiers)
                    IsIdentifiersToBeApplied = true;

                DelimiterAfterNoun = model.DelimiterAfterNoun ?? ", ";
                DelimiterAfterModifier = model.DelimiterAfterModifier ?? ", ";

                if (model.TruncationType == "M")
                    TruncationType = 'M';
                else
                    TruncationType = 'B';

                CharacterLimit = model.CharacterLimit;
                DelimiterForTruncation = model.DelimiterForTruncation != null ? Convert.ToChar(model.DelimiterForTruncation.Trim()) : ';';

                if (model.ThirdOrderOfDataInDescription.Trim().ToString().ToUpper() == "ATTRIBUTE" || model.ThirdOrderOfDataInDescription.Trim().ToString().ToUpper() == "A")
                    ThirdOrderOfDataInDescription = "A";

                if (model.ThirdOrderOfDataInDescription.Trim().ToString().ToUpper() == "ADDITIONAL INFO" || model.ThirdOrderOfDataInDescription.Trim().ToString().ToUpper() == "D")
                    ThirdOrderOfDataInDescription = "D";

                if (model.ThirdOrderOfDataInDescription.Trim().ToString().ToUpper() == "MFR DETAIL" || model.ThirdOrderOfDataInDescription.Trim().ToString().ToUpper() == "M")
                    ThirdOrderOfDataInDescription = "M";

                if (model.FourthOrderOfDataInDescription.Trim().ToString().ToUpper() == "ATTRIBUTE" || model.FourthOrderOfDataInDescription.Trim().ToString().ToUpper() == "A")
                    FourthOrderOfDataInDescription = "A";

                if (model.FourthOrderOfDataInDescription.Trim().ToString().ToUpper() == "ADDITIONAL INFO" || model.FourthOrderOfDataInDescription.Trim().ToString().ToUpper() == "D")
                    FourthOrderOfDataInDescription = "D";

                if (model.FourthOrderOfDataInDescription.Trim().ToString().ToUpper() == "MFR DETAIL" || model.FourthOrderOfDataInDescription.Trim().ToString().ToUpper() == "M")
                    FourthOrderOfDataInDescription = "M";

                if (model.FifthOrderOfDataInDescription.Trim().ToString().ToUpper() == "ATTRIBUTE" || model.FifthOrderOfDataInDescription.Trim().ToString().ToUpper() == "A")
                    FifthOrderOfDataInDescription = "A";

                if (model.FifthOrderOfDataInDescription.Trim().ToString().ToUpper() == "ADDITIONAL INFO" || model.FifthOrderOfDataInDescription.Trim().ToString().ToUpper() == "D")
                    FifthOrderOfDataInDescription = "D";

                if (model.FifthOrderOfDataInDescription.Trim().ToString().ToUpper() == "MFR DETAIL" || model.FifthOrderOfDataInDescription.Trim().ToString().ToUpper() == "M")
                    FifthOrderOfDataInDescription = "M";
                #endregion

                #region Opening Input File
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.LoadData(AttributeFilePath);

                Worksheet wsIF = wbIF.Worksheets[0];
                int maxRows = wsIF.Cells.MaxRow;
                int maxColumns = wsIF.Cells.MaxColumn;
                #endregion

                #region Adding Abbreviation File contents to Abbreviation List
                List<clsAbbreviation> AbbreviationList = new List<clsAbbreviation>();
                if (IsNounToBeAbbreviated || IsModifierToBeAbbreviated || IsAttributeNameToBeAbbreviated ||
                    IsAttributeValueToBeAbbreviated || IsAdditionalInformationToBeAbbreviated || IsMFRNameToBeAbbreviated)
                {
                    AbbreviationFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + model.AbbreviationFileName);
                    Workbook wbAF = new Workbook();
                    wbAF.LoadData(AbbreviationFilePath);
                    Worksheet wsAF = wbAF.Worksheets[0];
                    for (int row = 1; row <= wsAF.Cells.MaxRow; row++)
                    {
                        AbbreviationList.Add(new clsAbbreviation
                        {
                            ExpandedVersion = wsAF.Cells[row, 0].StringValue,
                            Abbreviation = wsAF.Cells[row, 1].StringValue
                        });
                    }
                    AbbreviationList = AbbreviationList.OrderByDescending(o => o.Abbreviation.Length).ToList();
                }
                #endregion

                #region Adding Identifier File contents to Identifier List
                List<Identifier> IdentifierList = new List<Identifier>();
                if (IsIdentifiersToBeApplied)
                {
                    IdentifierFilePath = HttpContext.Current.Server.MapPath(@"\temp\" + model.IdentifierFileName);
                    Workbook wbID = new Workbook();
                    wbID.LoadData(IdentifierFilePath);
                    Worksheet wsID = wbID.Worksheets[0];
                    for (int row = 1; row <= wsID.Cells.MaxRow; row++)
                    {
                        IdentifierList.Add(new Identifier
                        {
                            Characteristics = wsID.Cells[row, 0].StringValue,
                            Abbreviation = wsID.Cells[row, 1].StringValue,
                            SuffixORPrefix = wsID.Cells[row, 2].StringValue
                        });
                    }
                }
                #endregion

                #region Finding out column No. for Noun,Modifier,MFR Name1,MFR Part No1,Attribute Name1,Additional Information, Description
                for (int col = 0; col <= maxColumns; col++)
                {
                    if (wsIF.Cells[0, col].StringValue.Trim().ToUpper() == "NOUN")
                        NounColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToUpper() == "MODIFIER")
                        ModifierColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToUpper() == "MFR NAME1")
                        MFRName1ColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToUpper() == "ATTRIBUTE1")
                        AttributeName1ColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToUpper() == "ADDITIONAL INFO")
                        AdditionalInfoColumnNo = col;

                    if (wsIF.Cells[0, col].StringValue.Trim().ToUpper() == "DESCRIPTION")
                        DescriptionColumnNo = col;

                    if (NounColumnNo >= 0 && ModifierColumnNo >= 0 && MFRName1ColumnNo >= 0 && AttributeName1ColumnNo >= 0 && AdditionalInfoColumnNo >= 0 && DescriptionColumnNo >= 0)
                        break;
                }
                #endregion

                #region Processing Rows starts
                for (int row = 1; row <= maxRows; row++)
                {
                    Noun = string.Empty; Modifier = string.Empty; Description = string.Empty; LeftOutInformation = string.Empty;

                    #region Add First order of data to description
                    if (!IsNounExcluded)
                    {
                        Noun = wsIF.Cells[row, NounColumnNo].StringValue;
                        if (IsNounToBeAbbreviated)
                        {
                            if (!string.IsNullOrEmpty(Noun.Trim()))
                                Noun = GetAbbreviatedString(Noun, AbbreviationList);
                        }

                        if (!string.IsNullOrEmpty(Noun))
                            Description = Noun + DelimiterAfterNoun;
                    }
                    #endregion

                    #region Add Second order of data in description
                    if (!IsModifierExcluded)
                    {
                        Modifier = wsIF.Cells[row, ModifierColumnNo].StringValue;
                        if (model.SpecificModifierExcluded.Length > 0)
                        {
                            if (model.SpecificModifierExcluded.ToUpper() == Modifier.ToUpper())
                                Modifier = "";
                        }

                        if (!string.IsNullOrEmpty(Modifier))
                        {
                            if (IsModifierToBeAbbreviated)
                                Modifier = GetAbbreviatedString(Modifier, AbbreviationList);
                            Description = Description + Modifier + DelimiterAfterModifier;
                        }
                    }
                    #endregion

                    #region Add Third order of data in description
                    if (ThirdOrderOfDataInDescription == "M")
                    {
                        string[] returnArray = AddMFRDetailsToDescription(wsIF, row, Description, LeftOutInformation, MFRName1ColumnNo, IsMFRNameToBeAbbreviated, AbbreviationList, model);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    else if (ThirdOrderOfDataInDescription == "A")
                    {
                        string[] returnArray = AddAttributeDetailsToDescription(wsIF, row, Description, LeftOutInformation, AttributeName1ColumnNo, IsAttributeNameToBeAbbreviated, IsAttributeValueToBeAbbreviated, AbbreviationList, IsIdentifiersToBeApplied, IdentifierList, model);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    else if (ThirdOrderOfDataInDescription == "D")
                    {
                        string[] returnArray = AddAdditionalInformationToDescription(wsIF, model, row, Description, LeftOutInformation, AdditionalInfoColumnNo, IsAttributeNameToBeAbbreviated, IsAttributeValueToBeAbbreviated, IsAdditionalInformationToBeAbbreviated, AbbreviationList, IsIdentifiersToBeApplied, IdentifierList);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    #endregion

                    #region Add Fourth order of data in description
                    if (FourthOrderOfDataInDescription == "M")
                    {
                        string[] returnArray = AddMFRDetailsToDescription(wsIF, row, Description, LeftOutInformation, MFRName1ColumnNo, IsMFRNameToBeAbbreviated, AbbreviationList, model);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    else if (FourthOrderOfDataInDescription == "A")
                    {
                        string[] returnArray = AddAttributeDetailsToDescription(wsIF, row, Description, LeftOutInformation, AttributeName1ColumnNo, IsAttributeNameToBeAbbreviated, IsAttributeValueToBeAbbreviated, AbbreviationList, IsIdentifiersToBeApplied, IdentifierList, model);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    else if (FourthOrderOfDataInDescription == "D")
                    {
                        string[] returnArray = AddAdditionalInformationToDescription(wsIF, model, row, Description, LeftOutInformation, AdditionalInfoColumnNo, IsAttributeNameToBeAbbreviated, IsAttributeValueToBeAbbreviated, IsAdditionalInformationToBeAbbreviated, AbbreviationList, IsIdentifiersToBeApplied, IdentifierList);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    #endregion

                    #region Add Fifth order of data in description
                    if (FifthOrderOfDataInDescription == "M")
                    {
                        string[] returnArray = AddMFRDetailsToDescription(wsIF, row, Description, LeftOutInformation, MFRName1ColumnNo, IsMFRNameToBeAbbreviated, AbbreviationList, model);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    else if (FifthOrderOfDataInDescription == "A")
                    {
                        string[] returnArray = AddAttributeDetailsToDescription(wsIF, row, Description, LeftOutInformation, AttributeName1ColumnNo, IsAttributeNameToBeAbbreviated, IsAttributeValueToBeAbbreviated, AbbreviationList, IsIdentifiersToBeApplied, IdentifierList, model);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    else if (FifthOrderOfDataInDescription == "D")
                    {
                        string[] returnArray = AddAdditionalInformationToDescription(wsIF, model, row, Description, LeftOutInformation, AdditionalInfoColumnNo, IsAttributeNameToBeAbbreviated, IsAttributeValueToBeAbbreviated, IsAdditionalInformationToBeAbbreviated, AbbreviationList, IsIdentifiersToBeApplied, IdentifierList);
                        Description = returnArray[0];
                        LeftOutInformation = returnArray[1];
                    }
                    #endregion

                    #region Cut at character limit and decide meaningful (based on Trunction Character) or blind truncation (cut at exact limit)
                    if (model.DescriptionToGenerate == "S")
                    {
                        #region Removing the delimiter at the end of description for rows fitting description into character limit
                        if (Description.Length > 0)
                        {
                            if (Description.Length <= CharacterLimit + 2 && TruncationType == 'M')       //Checking next two characters after character limit are delimiter characters
                            {
                                if (!IsNounExcluded && Description.LastIndexOf(DelimiterAfterNoun) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterNoun), DelimiterAfterNoun.Length).ToUpper() == DelimiterAfterNoun.ToUpper() &&
                                    Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterNoun)).Length + DelimiterAfterNoun.Length)
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterNoun.Length, DelimiterAfterNoun.Length, string.Empty);
                                }

                                if (!IsModifierExcluded && Description.LastIndexOf(DelimiterAfterModifier) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterModifier), DelimiterAfterModifier.Length).ToUpper() == DelimiterAfterModifier.ToUpper() &&
                                        Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterModifier)).Length + DelimiterAfterModifier.Length)
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterModifier.Length, DelimiterAfterModifier.Length, string.Empty);
                                }

                                if (!IsMFRNameExcluded && Description.LastIndexOf(DelimiterAfterMFRName) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterMFRName), DelimiterAfterMFRName.Length).ToUpper() == DelimiterAfterMFRName.ToUpper() &&
                                        Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterMFRName)).Length + DelimiterAfterMFRName.Length)
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterMFRName.Length, DelimiterAfterMFRName.Length, string.Empty);
                                }

                                if (!IsMFRPartNoExcluded && Description.LastIndexOf(DelimiterAfterMFRPartNo) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterMFRPartNo), DelimiterAfterMFRPartNo.Length).ToUpper() == DelimiterAfterMFRPartNo.ToUpper() &&
                                        Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterMFRPartNo)).Length + DelimiterAfterMFRPartNo.Length)
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterMFRPartNo.Length, DelimiterAfterMFRPartNo.Length, string.Empty);
                                }

                                if (!IsAttributeNameExcluded && Description.LastIndexOf(DelimiterAfterAttributeName) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterAttributeName), DelimiterAfterAttributeName.Length).ToUpper() == DelimiterAfterAttributeName.ToUpper() &&
                                        Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAttributeName)).Length + DelimiterAfterAttributeName.Length)
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterAttributeName.Length, DelimiterAfterAttributeName.Length, string.Empty);
                                }

                                if (!IsAttributeValueExcluded && Description.LastIndexOf(DelimiterAfterAttributeValue) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterAttributeValue), DelimiterAfterAttributeValue.Length).ToUpper() == DelimiterAfterAttributeValue.ToUpper() &&
                                        Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAttributeValue)).Length + DelimiterAfterAttributeValue.Length)
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterAttributeValue.Length, DelimiterAfterAttributeValue.Length, string.Empty);
                                }

                                if (!IsAdditionalInformationExcluded && Description.LastIndexOf(DelimiterAfterAdditionalInformation) > -1)
                                {
                                    if (Description.Substring(Description.LastIndexOf(DelimiterAfterAdditionalInformation), DelimiterAfterAdditionalInformation.Length).ToUpper() == DelimiterAfterAdditionalInformation.ToUpper() &&
                                        Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAdditionalInformation)).Length + DelimiterAfterAdditionalInformation.Length)
                                        Description = ReplaceAt(Description, Description.Length - DelimiterAfterAdditionalInformation.Length, DelimiterAfterAdditionalInformation.Length, string.Empty);
                                }
                            }
                        }
                        #endregion

                        #region Cut blindly at character limit (only for description length exceeding character limit) and push to left out information. If meaningful, truncate till truncation character and push to left out information
                        if (Description.Length > CharacterLimit)
                        {
                            LeftOutInformation = Description.Substring(CharacterLimit, Description.Length - CharacterLimit) + LeftOutInformation;
                            Description = Description.Substring(0, CharacterLimit);

                            if (TruncationType == 'M')
                            {
                                if (Description.IndexOf(DelimiterForTruncation) > -1)
                                {
                                    LeftOutInformation = Description.Substring(Description.LastIndexOf(DelimiterForTruncation), Description.Length - Description.LastIndexOf(DelimiterForTruncation)) + LeftOutInformation;
                                    Description = Description.Substring(0, Description.LastIndexOf(DelimiterForTruncation));
                                }
                                else
                                {
                                    LeftOutInformation = Description + LeftOutInformation;
                                    Description = string.Empty;
                                }
                            }
                        }
                        #endregion

                        #region Remove the characters at the beginning and end of left out information
                        if (LeftOutInformation.Trim().Length > 0)
                        {
                            #region Removing the delimited characters of left out at the begining
                            if (!IsNounExcluded)
                            {
                                if (LeftOutInformation.Substring(0, DelimiterAfterNoun.Length) == DelimiterAfterNoun)
                                    LeftOutInformation = LeftOutInformation.Substring(DelimiterAfterNoun.Length, LeftOutInformation.Length - DelimiterAfterNoun.Length);
                            }

                            if (!IsModifierExcluded)
                            {
                                if (LeftOutInformation.Substring(0, DelimiterAfterModifier.Length) == DelimiterAfterModifier)
                                    LeftOutInformation = LeftOutInformation.Substring(DelimiterAfterModifier.Length, LeftOutInformation.Length - DelimiterAfterModifier.Length);
                            }

                            if (!IsMFRNameExcluded)
                            {
                                if (LeftOutInformation.Substring(0, DelimiterAfterMFRName.Length) == DelimiterAfterMFRName)
                                    LeftOutInformation = LeftOutInformation.Substring(DelimiterAfterMFRName.Length, LeftOutInformation.Length - DelimiterAfterMFRName.Length);
                            }

                            if (!IsMFRPartNoExcluded)
                            {
                                if (LeftOutInformation.Substring(0, DelimiterAfterMFRPartNo.Length) == DelimiterAfterMFRPartNo)
                                    LeftOutInformation = LeftOutInformation.Substring(DelimiterAfterMFRPartNo.Length, LeftOutInformation.Length - DelimiterAfterMFRPartNo.Length);
                            }

                            if (!IsAttributeNameExcluded)
                            {
                                if (LeftOutInformation.Substring(0, DelimiterAfterAttributeName.Length) == DelimiterAfterAttributeName)
                                    LeftOutInformation = LeftOutInformation.Substring(DelimiterAfterAttributeName.Length, LeftOutInformation.Length - DelimiterAfterAttributeName.Length);
                            }

                            if (!IsAttributeValueExcluded)
                            {
                                if (LeftOutInformation.Substring(0, DelimiterAfterAttributeValue.Length) == DelimiterAfterAttributeValue)
                                    LeftOutInformation = LeftOutInformation.Substring(DelimiterAfterAttributeValue.Length, LeftOutInformation.Length - DelimiterAfterAttributeValue.Length);

                                if (LeftOutInformation.Substring(0, 1) == MultipleAttributeValuesSeparator.ToString())
                                    LeftOutInformation = LeftOutInformation.Substring(1, LeftOutInformation.Length - 1);
                            }

                            if (!IsAdditionalInformationExcluded)
                            {
                                if (LeftOutInformation.Substring(0, DelimiterAfterAdditionalInformation.Length) == DelimiterAfterAdditionalInformation)
                                    LeftOutInformation = LeftOutInformation.Substring(DelimiterAfterAdditionalInformation.Length, LeftOutInformation.Length - DelimiterAfterAdditionalInformation.Length);
                            }
                            #endregion

                            #region Removing the last delimited characters of left out
                            if (!IsNounExcluded)
                            {
                                if (LeftOutInformation.Substring(LeftOutInformation.Length - DelimiterAfterNoun.Length, DelimiterAfterNoun.Length) == DelimiterAfterNoun)
                                    LeftOutInformation = LeftOutInformation.Substring(0, LeftOutInformation.Length - DelimiterAfterNoun.Length);
                            }

                            if (!IsModifierExcluded)
                            {
                                if (LeftOutInformation.Substring(LeftOutInformation.Length - DelimiterAfterModifier.Length, DelimiterAfterModifier.Length) == DelimiterAfterModifier)
                                    LeftOutInformation = LeftOutInformation.Substring(0, LeftOutInformation.Length - DelimiterAfterModifier.Length);
                            }

                            if (!IsMFRNameExcluded)
                            {
                                if (LeftOutInformation.Substring(LeftOutInformation.Length - DelimiterAfterMFRName.Length, DelimiterAfterMFRName.Length) == DelimiterAfterMFRName)
                                    LeftOutInformation = LeftOutInformation.Substring(0, LeftOutInformation.Length - DelimiterAfterMFRName.Length);
                            }

                            if (!IsMFRPartNoExcluded)
                            {
                                if (LeftOutInformation.Substring(LeftOutInformation.Length - DelimiterAfterMFRPartNo.Length, DelimiterAfterMFRPartNo.Length) == DelimiterAfterMFRPartNo)
                                    LeftOutInformation = LeftOutInformation.Substring(0, LeftOutInformation.Length - DelimiterAfterMFRPartNo.Length);
                            }

                            if (!IsAttributeNameExcluded)
                            {
                                if (LeftOutInformation.Substring(LeftOutInformation.Length - DelimiterAfterAttributeName.Length, DelimiterAfterAttributeName.Length) == DelimiterAfterAttributeName)
                                    LeftOutInformation = LeftOutInformation.Substring(0, LeftOutInformation.Length - DelimiterAfterAttributeName.Length);
                            }

                            if (!IsAttributeValueExcluded)
                            {
                                if (LeftOutInformation.Substring(LeftOutInformation.Length - DelimiterAfterAttributeValue.Length, DelimiterAfterAttributeValue.Length) == DelimiterAfterAttributeValue)
                                    LeftOutInformation = LeftOutInformation.Substring(0, LeftOutInformation.Length - DelimiterAfterAttributeValue.Length);
                            }

                            if (!IsAdditionalInformationExcluded)
                            {
                                if (LeftOutInformation.Substring(LeftOutInformation.Length - DelimiterAfterAdditionalInformation.Length, DelimiterAfterAdditionalInformation.Length) == DelimiterAfterAdditionalInformation)
                                    LeftOutInformation = LeftOutInformation.Substring(0, LeftOutInformation.Length - DelimiterAfterAdditionalInformation.Length);
                            }
                            #endregion
                        }
                        #endregion

                        wsIF.Cells[row, DescriptionColumnNo].PutValue(Description.Trim());
                        wsIF.Cells[row, DescriptionColumnNo + 1].PutValue(LeftOutInformation.Trim());
                    }
                    else
                    {
                        #region Remove the last delimited characters of Description
                        if (!IsNounExcluded && Description.LastIndexOf(DelimiterAfterNoun) > -1)
                        {
                            if (Description.Substring(Description.LastIndexOf(DelimiterAfterNoun), DelimiterAfterNoun.Length).ToUpper() == DelimiterAfterNoun.ToUpper() &&
                                Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterNoun)).Length + DelimiterAfterNoun.Length)
                            {
                                Description = ReplaceAt(Description, Description.Length - DelimiterAfterNoun.Length, DelimiterAfterNoun.Length, string.Empty);
                            }
                        }

                        if (!IsModifierExcluded && Description.LastIndexOf(DelimiterAfterModifier) > -1)
                        {
                            if (Description.Substring(Description.LastIndexOf(DelimiterAfterModifier), DelimiterAfterModifier.Length).ToUpper() == DelimiterAfterModifier.ToUpper() &&
                                Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterModifier)).Length + DelimiterAfterModifier.Length)
                            {
                                Description = ReplaceAt(Description, Description.Length - DelimiterAfterModifier.Length, DelimiterAfterModifier.Length, string.Empty);
                            }
                        }

                        if (!IsMFRNameExcluded && Description.LastIndexOf(DelimiterAfterMFRName) > -1)
                        {
                            if (Description.Substring(Description.LastIndexOf(DelimiterAfterMFRName), DelimiterAfterMFRName.Length).ToUpper() == DelimiterAfterMFRName.ToUpper() &&
                                Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterMFRName)).Length + DelimiterAfterMFRName.Length)
                            {
                                Description = ReplaceAt(Description, Description.Length - DelimiterAfterMFRName.Length, DelimiterAfterMFRName.Length, string.Empty);
                            }
                        }

                        if (!IsMFRPartNoExcluded && Description.LastIndexOf(DelimiterAfterMFRPartNo) > -1)
                        {
                            if (Description.Substring(Description.LastIndexOf(DelimiterAfterMFRPartNo), DelimiterAfterMFRPartNo.Length).ToUpper() == DelimiterAfterMFRPartNo.ToUpper() &&
                                Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterMFRPartNo)).Length + DelimiterAfterMFRPartNo.Length)
                            {
                                Description = ReplaceAt(Description, Description.Length - DelimiterAfterMFRPartNo.Length, DelimiterAfterMFRPartNo.Length, string.Empty);
                            }
                        }

                        if (!IsAttributeNameExcluded && Description.LastIndexOf(DelimiterAfterAttributeName) > -1)
                        {
                            if (Description.Substring(Description.LastIndexOf(DelimiterAfterAttributeName), DelimiterAfterAttributeName.Length).ToUpper() == DelimiterAfterAttributeName.ToUpper() &&
                                Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAttributeName)).Length + DelimiterAfterAttributeName.Length)
                            {
                                Description = ReplaceAt(Description, Description.Length - DelimiterAfterAttributeName.Length, DelimiterAfterAttributeName.Length, string.Empty);
                            }
                        }

                        if (!IsAttributeValueExcluded && Description.LastIndexOf(DelimiterAfterAttributeValue) > -1)
                        {
                            if (Description.Substring(Description.LastIndexOf(DelimiterAfterAttributeValue), DelimiterAfterAttributeValue.Length).ToUpper() == DelimiterAfterAttributeValue.ToUpper() &&
                                Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAttributeValue)).Length + DelimiterAfterAttributeValue.Length)
                            {
                                Description = ReplaceAt(Description, Description.Length - DelimiterAfterAttributeValue.Length, DelimiterAfterAttributeValue.Length, string.Empty);
                            }
                        }

                        if (!IsAdditionalInformationExcluded && Description.LastIndexOf(DelimiterAfterAdditionalInformation) > -1)
                        {
                            if (Description.Substring(Description.LastIndexOf(DelimiterAfterAdditionalInformation), DelimiterAfterAdditionalInformation.Length).ToUpper() == DelimiterAfterAdditionalInformation.ToUpper() &&
                                Description.Length == Description.Substring(0, Description.LastIndexOf(DelimiterAfterAdditionalInformation)).Length + DelimiterAfterAdditionalInformation.Length)
                            {
                                Description = ReplaceAt(Description, Description.Length - DelimiterAfterAdditionalInformation.Length, DelimiterAfterAdditionalInformation.Length, string.Empty);
                            }
                        }
                        #endregion

                        wsIF.Cells[row, DescriptionColumnNo].PutValue(Description.Trim().Replace("  ", " "));
                    }
                    #endregion
                }

                #region Save and Download the file
                wsIF.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + model.InputFileName;
                wbIF.Save(filename);
                #endregion

                #endregion

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
