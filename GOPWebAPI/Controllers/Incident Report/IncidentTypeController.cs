using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.Http.ExceptionHandling;
using System.Web.Http.Results;
using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.Incident_Report_Models;
using Newtonsoft.Json.Linq;

namespace GOPWebAPI.Controllers.Incident_Report
{
    [RoutePrefix("api/IncidentType")]
    public class IncidentTypeController : ApiController
    {
        private BLLIncidentType _BLLIncidentType;
        public IncidentTypeController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLIncidentType = new BLLIncidentType(connectionString);
        }

        #region Add Incident Type
        [HttpPost]
        [Route("PostAddIncidentType")]
        public HttpResponseMessage PostAddIncidentType([FromBody] IncidentTypeModel obj)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(obj.UserID, "Incident Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                Result objResult = new Result();

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Model Data");

                // Call the service to add Incident Type
                string strOutput = _BLLIncidentType.AddIncidentType(obj);

                if (!string.IsNullOrEmpty(strOutput) && strOutput.ToLower().StartsWith("success:"))
                {
                    objResult.Msg = strOutput;
                    objResult.Success = 1;
                    return Request.CreateResponse(HttpStatusCode.Created, objResult);
                }
                else
                {
                    objResult.Msg = strOutput;
                    objResult.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.Conflict, objResult);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Read All Incident Types
        [HttpGet]
        [Route("ReadAllIncidentTypes")]
        public HttpResponseMessage ReadAllIncidentTypes(string UserID, bool IsToFetchOnlyActiveIncidentTypes = true)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Incident Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                DataTable dtDetails = new DataTable();
                List<IncidentTypeModel> lstIncidentTypes = new List<IncidentTypeModel>();
                dtDetails = _BLLIncidentType.ReadAllIncidentTypes();

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentTypeModel objIncidentTypeModel = new IncidentTypeModel();
                        objIncidentTypeModel.IncidentTypeID = dtDetails.Rows[i]["IncidentTypeID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["IncidentTypeID"]) : 0;
                        objIncidentTypeModel.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        objIncidentTypeModel.IsActive = dtDetails.Rows[i]["IsActive"] != DBNull.Value ? Convert.ToBoolean(dtDetails.Rows[i]["IsActive"]) : false;
                        lstIncidentTypes.Add(objIncidentTypeModel);
                    }

                    if (IsToFetchOnlyActiveIncidentTypes)
                    {
                        JObject PEReport = new JObject(new JProperty("Success", 1),
                        new JProperty("RecordCount", dtDetails.Rows.Count),
                        new JProperty("IncidentTypes", new JArray(from p in lstIncidentTypes.Where(it => it.IsActive == true)
                                                                  select new JObject(
                                                                    new JProperty("IncidentTypeID", p.IncidentTypeID),
                                                                    new JProperty("IncidentType", p.IncidentType),
                                                                    new JProperty("IsActive", p.IsActive)
                                                                    ))));
                        return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                    }
                    else
                    {
                        JObject PEReport = new JObject(new JProperty("Success", 1),
                        new JProperty("RecordCount", dtDetails.Rows.Count),
                        new JProperty("IncidentTypes", new JArray(from p in lstIncidentTypes
                                                                  select new JObject(
                                                                    new JProperty("IncidentTypeID", p.IncidentTypeID),
                                                                    new JProperty("IncidentType", p.IncidentType),
                                                                    new JProperty("IsActive", p.IsActive)
                                                                    ))));
                        return Request.CreateResponse(HttpStatusCode.OK, PEReport);
                    }
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

        #region Read Incident Type By Id
        [HttpGet]
        [Route("ReadIncidentTypeById")]
        public HttpResponseMessage ReadIncidentTypeById(string UserID, int IncidentTypeID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Incident Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                DataTable dtDetails = new DataTable();
                dtDetails = _BLLIncidentType.ReadIncidentTypeById(IncidentTypeID);

                if (dtDetails.Rows.Count > 0)
                {
                    IncidentTypeModel objIncidentTypeModel = new IncidentTypeModel();
                    objIncidentTypeModel.IncidentTypeID = dtDetails.Rows[0]["IncidentTypeID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[0]["IncidentTypeID"]) : 0;
                    objIncidentTypeModel.IncidentType = dtDetails.Rows[0]["IncidentType"] != DBNull.Value ? dtDetails.Rows[0]["IncidentType"].ToString().Trim() : "";
                    objIncidentTypeModel.IsActive = dtDetails.Rows[0]["IsActive"] != DBNull.Value ? Convert.ToBoolean(dtDetails.Rows[0]["IsActive"]) : false;

                    JObject PEReport = new JObject(new JProperty("Success", 1),
                                                new JProperty("Incident Type", new JObject(
                                                                    new JProperty("IncidentTypeID", objIncidentTypeModel.IncidentTypeID),
                                                                    new JProperty("IncidentType", objIncidentTypeModel.IncidentType),
                                                                    new JProperty("IsActive", objIncidentTypeModel.IsActive))));

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

        #region Update Incident Type
        [HttpPost]
        [Route("PostUpdateIncidentType")]
        public HttpResponseMessage PostUpdateIncidentType([FromBody] IncidentTypeModel obj)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(obj.UserID, "Incident Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                Result objVar = new Result();

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Model Data");

                string strOutput = _BLLIncidentType.UpdateIncidentType(obj);

                if (!string.IsNullOrEmpty(strOutput) && strOutput.ToLower().StartsWith("success:"))
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 1;
                    return Request.CreateResponse(HttpStatusCode.Created, objVar);
                }
                else
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.Conflict, objVar);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Delete Incident Type
        [HttpPost]
        [Route("PostDeleteIncidentType")]
        public HttpResponseMessage PostDeleteIncidentType([FromBody] IncidentTypeModel obj)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(obj.UserID, "Incident Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                Result objVar = new Result();

                if (obj == null)
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "Incident Type data is required");

                string strOutput = _BLLIncidentType.DeleteIncidentType(obj);

                if (!string.IsNullOrEmpty(strOutput) && strOutput.ToLower().StartsWith("success:"))
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 1;
                    return Request.CreateResponse(HttpStatusCode.Created, objVar);
                }
                else
                {
                    objVar.Msg = strOutput;
                    objVar.Success = 0;
                    return Request.CreateResponse(HttpStatusCode.Conflict, objVar);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Export Incident Types List to Excel
        [HttpGet]
        [Route("ExportIncidentTypesListToExcel")]
        public HttpResponseMessage ExportIncidentTypesListToExcel(string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Incident Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                string FileName = "Incident Types List.xlsx";
                DataTable dtDetails = new DataTable();
                List<IncidentTypeModel> lstIncidentTypes = new List<IncidentTypeModel>();
                dtDetails = _BLLIncidentType.ReadAllIncidentTypes();

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        IncidentTypeModel objIncidentTypeModel = new IncidentTypeModel();
                        objIncidentTypeModel.IncidentTypeID = dtDetails.Rows[i]["IncidentTypeID"] != DBNull.Value ? Convert.ToInt32(dtDetails.Rows[i]["IncidentTypeID"]) : 0;
                        objIncidentTypeModel.IncidentType = dtDetails.Rows[i]["IncidentType"] != DBNull.Value ? dtDetails.Rows[i]["IncidentType"].ToString().Trim() : "";
                        objIncidentTypeModel.IsActive = dtDetails.Rows[i]["IsActive"] != DBNull.Value ? Convert.ToBoolean(dtDetails.Rows[i]["IsActive"]) : false;
                        lstIncidentTypes.Add(objIncidentTypeModel);
                    }

                    if (HttpContext.Current == null)
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);

                    //Create HTTP Response.
                    HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                    #region Setting up the workbook
                    Workbook wb = new Workbook();
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    wb.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                    var ws = wb.Worksheets[0];
                    int row = 1;
                    #endregion

                    #region Setting Styles
                    Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                    Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();

                    styleHeader.IsTextWrapped = true;
                    styleHeader.HorizontalAlignment = TextAlignmentType.Center;
                    styleHeader.VerticalAlignment = TextAlignmentType.Center;
                    styleHeader.ForegroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                    styleHeader.BackgroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                    styleHeader.Pattern = BackgroundType.VerticalStripe;
                    styleHeader.Font.Color = System.Drawing.Color.Black;
                    styleHeader.Font.IsBold = true;
                    styleHeader.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    styleHeader.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    styleHeader.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    styleHeader.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    styleHeader.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    styleCenterAlignData.HorizontalAlignment = TextAlignmentType.Center;
                    styleCenterAlignData.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    styleCenterAlignData.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    styleCenterAlignData.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    styleCenterAlignData.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    styleCenterAlignData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                    styleLeftAlignData.HorizontalAlignment = TextAlignmentType.Left;
                    styleLeftAlignData.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    styleLeftAlignData.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    styleLeftAlignData.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    styleLeftAlignData.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    styleLeftAlignData.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                    #endregion

                    #region Writing column headings and setting style
                    ws.Cells[0, 0].PutValue("S.No.");
                    ws.Cells[0, 1].PutValue("Incident Type ID");
                    ws.Cells[0, 2].PutValue("Incident Type");
                    ws.Cells[0, 3].PutValue("Is Active");

                    for (int c = 0; c <= 3; c++)
                        ws.Cells[0, c].SetStyle(styleHeader);
                    #endregion

                    #region Writing row data
                    foreach (IncidentTypeModel itmodel in lstIncidentTypes)
                    {
                        #region Writing row data
                        ws.Cells[row, 0].PutValue(row);
                        ws.Cells[row, 1].PutValue(itmodel.IncidentTypeID);
                        ws.Cells[row, 2].PutValue(itmodel.IncidentType);
                        ws.Cells[row, 3].PutValue(itmodel.IsActive==true?"Yes":"No");
                        #endregion

                        #region setting row data style
                        ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                        ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                        ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                        #endregion

                        row++;
                    }
                    #endregion

                    #region Saving and downloading the report
                    ws.AutoFitColumns();
                    string tempPath = System.IO.Path.GetTempPath();
                    string filename = tempPath + FileName;
                    wb.Save(filename);

                    //Read the file into a Byte Array.
                    byte[] bytes = File.ReadAllBytes(filename);

                    //Set the Response Content.
                    response.Content = new ByteArrayContent(bytes);

                    //Set the Response Content Length.
                    response.Content.Headers.ContentLength = bytes.LongLength;

                    //Set the Content Disposition Header Value and FileName.
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                    response.Content.Headers.ContentDisposition.FileName = FileName;

                    //Set the File Content Type.
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(filename));

                    return response;
                    #endregion
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
    }
}
