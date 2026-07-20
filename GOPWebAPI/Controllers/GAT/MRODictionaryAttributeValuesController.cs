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
    [RoutePrefix("api/MRODictionaryAttributeValues")]
    public class MRODictionaryAttributeValuesController : ApiController
    {
        private BLLGATMRODictionaryAttributeValues _BLLGATMRODictionaryAttributeValues;

        public MRODictionaryAttributeValuesController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLGATMRODictionaryAttributeValues = new BLLGATMRODictionaryAttributeValues(connectionString);
        }

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
                if (wsIF.Cells[0, 0].StringValue.Trim().ToUpper() != "SL.NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SL.No' not found in input file first worksheet. Please select a valid file.");
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

        #region Fetch Noun Modifier Attributes Values and Write to output file
        [HttpPost]
        [Route("FetchNounModifierAttributeValuesAndWriteToOutputFile")]
        public HttpResponseMessage FetchNounModifierAttributeValuesAndWriteToOutputFile(string UploadedInputFileName, string InputFileName, string VersionNameOrNo)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            try
            {
                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet wsIF = wbIF.Worksheets[0];

                #region Setting the styles
                Workbook wbGriha = new Workbook();
                wbGriha.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var wELReport = wbGriha.Worksheets[0];
                Aspose.Cells.Style styleNMHeader = wELReport.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleAttributeHeader = wELReport.Cells[0, 0].GetStyle();

                styleNMHeader.IsTextWrapped = true;
                styleNMHeader.HorizontalAlignment = TextAlignmentType.Center;
                styleNMHeader.VerticalAlignment = TextAlignmentType.Center;
                styleNMHeader.ForegroundColor = System.Drawing.Color.FromArgb(0, 231, 230, 230);
                styleNMHeader.BackgroundColor = System.Drawing.Color.FromArgb(0, 231, 230, 230);
                styleNMHeader.Pattern = BackgroundType.VerticalStripe;
                styleNMHeader.Font.Color = System.Drawing.Color.Black;
                styleNMHeader.Font.IsBold = true;
                styleNMHeader.Font.Size = 11;

                styleAttributeHeader.IsTextWrapped = true;
                styleAttributeHeader.HorizontalAlignment = TextAlignmentType.Center;
                styleAttributeHeader.VerticalAlignment = TextAlignmentType.Center;
                styleAttributeHeader.ForegroundColor = System.Drawing.Color.FromArgb(0, 169, 208, 142);
                styleAttributeHeader.BackgroundColor = System.Drawing.Color.FromArgb(0, 169, 208, 142);
                styleAttributeHeader.Pattern = BackgroundType.VerticalStripe;
                styleAttributeHeader.Font.Color = System.Drawing.Color.Black;
                styleAttributeHeader.Font.IsBold = true;
                styleAttributeHeader.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                styleAttributeHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                styleAttributeHeader.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                styleAttributeHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                styleAttributeHeader.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                styleAttributeHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                styleAttributeHeader.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                styleAttributeHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                styleAttributeHeader.Font.Size = 11;
                #endregion

                DataTable dataTable = _BLLGATMRODictionaryAttributeValues.FetchMRODictionaryAttributeValues(VersionNameOrNo);

                List<NMAV> NMAttributeValueList= new List<NMAV>();
                foreach (DataRow dr in dataTable.Rows)
                {
                    NMAttributeValueList.Add(new NMAV
                    {
                        Noun = dr["Noun"].ToString(),
                        Modifier = dr["Modifier"].ToString(),
                        Attribute = dr["Attribute"].ToString(),
                        Value = dr["Value"].ToString(),
                    });
                }

                string Noun = string.Empty, Modifier = string.Empty, Attribute = string.Empty;
                for (int iRow = 1; iRow <= wsIF.Cells.MaxRow; iRow++)             //Start Iterating from first to last row of Input Worksheet
                {
                    Noun = wsIF.Cells[iRow, 1].StringValue.Trim();
                    Modifier = wsIF.Cells[iRow, 2].StringValue.Trim();

                    #region Add Worksheet and get index of it
                    wbIF.Worksheets.Add(Noun + "_" + Modifier.Replace("/", "or"));
                    int wsIdx = -1;
                    foreach (Worksheet s in wbIF.Worksheets)
                    {
                        if (s.Name == Noun + "_" + Modifier.Replace("/", "or"))
                        {
                            wsIdx++;
                            break;
                        }
                        else
                            wsIdx++;
                    }
                    #endregion

                    if (wsIdx > -1)
                    {
                        Worksheet wsa = wbIF.Worksheets[wsIdx];

                        wsa.Cells[0, 0].PutValue(Noun);
                        wsa.Cells[0, 0].SetStyle(styleNMHeader);
                        wsa.Cells[0, 1].PutValue(Modifier);
                        wsa.Cells[0, 1].SetStyle(styleNMHeader);
                        wsa.Cells.SetRowHeight(0, 15);

                        #region Get Distinct Attributes of Noun-Modifier
                        var NMAttributeList = NMAttributeValueList.Where(al => al.Noun == Noun && al.Modifier == Modifier).Select(a => new
                        {
                            Attribute = a.Attribute
                        }).Distinct().ToList();
                        #endregion

                        #region Write All Attribute Names of Noun-Modifier
                        int col = 2;
                        for (int idx = 0; idx < NMAttributeList.Count(); idx++)
                        {
                            wsa.Cells[0, col].PutValue(NMAttributeList[idx].Attribute);
                            wsa.Cells[0, col].SetStyle(styleAttributeHeader);
                            col++;
                        }
                        #endregion

                        #region Write Each Attribute Values
                        for (int c = 2; c < col; c++)
                        {
                            Attribute = wsa.Cells[0, c].StringValue.Trim().ToUpper();
                            if (!string.IsNullOrEmpty(Attribute))
                            {
                                var AttributeValueList = NMAttributeValueList.Where(vl => vl.Noun == Noun && vl.Modifier == Modifier && vl.Attribute == Attribute).Select(v => new
                                {
                                    Value = v.Value
                                }).ToList();

                                int vr = 1;
                                for (int vix = 0; vix < AttributeValueList.Count(); vix++)
                                {
                                    wsa.Cells[vr, c].PutValue(AttributeValueList[vix].Value);
                                    wsa.Cells.SetRowHeight(vr, 15);
                                    vr++;
                                }
                            }
                        }
                        #endregion

                        wsa.AutoFitColumns();
                    }
                }

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
    }
}
