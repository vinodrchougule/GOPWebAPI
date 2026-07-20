using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.Grievance_Reach_CEO_Directly_Models;

namespace GOPWebAPI.Controllers.Grievance_Reach_CEO_Directly
{
    [RoutePrefix("api/GrievanceReachCEODirectly")]
    public class GrievanceReachCEODirectlyController : ApiController
    {
        #region Create Suggestion
        [HttpPost]
        [Route("CreateSuggestion")]
        public HttpResponseMessage CreateSuggestion([FromBody] GrievanceReachCEODirectlyModel model)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid model Data");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spSuggestionToManagement";

                    //Add parameters with values - Mode - 1 - Create
                    cmd.Parameters.AddWithValue("@Subject", model.Subject);
                    cmd.Parameters.AddWithValue("@Details", model.Details);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Create);

                    //Calling sp to create suggestion
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.ToLower() == "created")
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Suggestions By Status
        [HttpGet]
        [Route("ReadSuggestionsByStatus")]
        public IHttpActionResult ReadSuggestionsByStatus(string UserID, string Status="O")
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Suggestions"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of suggestions
                List<GrievanceReachCEODirectlyModel> SuggestionsList = new List<GrievanceReachCEODirectlyModel>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSuggestionToManagement";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 6 - Read By Status
                    cmd.Parameters.AddWithValue("@Status", Status);
                    cmd.Parameters.AddWithValue("@Mode", 6);

                    //Calling sp to get list of suggestions
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        GrievanceReachCEODirectlyModel model = new GrievanceReachCEODirectlyModel();
                        model.SuggestionToManagementID = Convert.ToInt64(sqlReader["SuggestionToManagementID"]);
                        model.Subject = sqlReader["Subject"].ToString();
                        model.Details = sqlReader["Details"].ToString();
                        model.CreatedOn = sqlReader["CreatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["CreatedOn"];
                        SuggestionsList.Add(model);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(SuggestionsList);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Suggestion By Id
        [HttpGet]
        [Route("ReadSuggestionById")]
        public IHttpActionResult ReadSuggestionById(long id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Suggestions"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a suggestion instance
                GrievanceReachCEODirectlyModel model = new GrievanceReachCEODirectlyModel();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSuggestionToManagement";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@SuggestionToManagementID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get suggestion details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        model.SuggestionToManagementID = Convert.ToInt32(sqlReader["SuggestionToManagementID"]);
                        model.Subject = sqlReader["Subject"].ToString();
                        model.Details = sqlReader["Details"].ToString();
                        model.CreatedOn = sqlReader["CreatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["CreatedOn"];
                        conn.Close();

                        //return suggestion to the request
                        return Ok(model);
                    }
                    else
                    {
                        conn.Close();
                        return NotFound();
                    }
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Update Suggestion Status To Closed
        [HttpPatch]
        [Route("UpdateSuggestionStatusToClosed")]
        public HttpResponseMessage UpdateSuggestionStatusToClosed(long id,string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Update Suggestion"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spSuggestionToManagement";

                    //Add parameters with values - Mode - 2 - Update
                    cmd.Parameters.AddWithValue("@SuggestionToManagementID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);

                    //Calling sp to update suggestion
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.ToLower() == "updated")
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                }
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }
        }
        #endregion

        #region Download Suggestions to Excel
        [HttpGet]
        [Route("ExportSuggestionsToExcel")]
        public HttpResponseMessage ExportSuggestionsToExcel(string UserID, string Status = "O")
        {
            try
            {
                int SlNo = 1;
                string FileName = "Suggestions To Management.xlsx";

                if (!AccessControl.CanUserAccessPage(UserID, "View Suggestions"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the suggestions Data
                List<GrievanceReachCEODirectlyModel> SuggestionsList = new List<GrievanceReachCEODirectlyModel>();

                System.Data.Common.DbDataReader sqlReader;
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
                int row = 4;
                #endregion

                #region Setting Styles
                Aspose.Cells.Style styleHeader = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleCenterAlignData = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleLeftAlignData = ws.Cells[0, 0].GetStyle();
                Aspose.Cells.Style styleRightAlignData = ws.Cells[0, 0].GetStyle();

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

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSuggestionToManagement";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 6 - Read By Status
                    cmd.Parameters.AddWithValue("@Status", Status);
                    cmd.Parameters.AddWithValue("@Mode", 6);

                    //Calling sp to get list of suggestions
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        #region Adding data to List
                        while (sqlReader.Read())
                        {
                            GrievanceReachCEODirectlyModel model = new GrievanceReachCEODirectlyModel();
                            model.SuggestionToManagementID = Convert.ToInt64(sqlReader["SuggestionToManagementID"]);
                            model.Subject = sqlReader["Subject"].ToString();
                            model.Details = sqlReader["Details"].ToString();
                            model.CreatedOn = sqlReader["CreatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["CreatedOn"];
                            SuggestionsList.Add(model);
                        }
                        conn.Close();
                        #endregion

                        #region Writing report header
                        for (int col = 0; col <= 3; col++)
                            ws.Cells[1, col].SetStyle(styleCenterAlignData);

                        ws.Cells.Merge(1, 0, 1, 4);
                        ws.Cells[1, 0].PutValue("Suggestions To Management");
                        #endregion

                        #region Writing column headings
                        ws.Cells[3, 0].PutValue("S.No.");
                        ws.Cells[3, 1].PutValue("Subject");
                        ws.Cells[3, 2].PutValue("Details");
                        ws.Cells[3, 3].PutValue("Date");

                        for (int c = 0; c <= 3; c++)
                            ws.Cells[3, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (GrievanceReachCEODirectlyModel model in SuggestionsList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(SlNo);
                            ws.Cells[row, 1].PutValue(model.Subject);
                            ws.Cells[row, 2].PutValue(model.Details);
                            if (model.CreatedOn != null)
                                ws.Cells[row, 3].PutValue(Convert.ToDateTime(model.CreatedOn).ToString("dd-MMM-yyyy"));
                            #endregion

                            #region setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                            #endregion
                            
                            row++; SlNo++;
                        }
                        #endregion
                    }

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
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion
    }
}
