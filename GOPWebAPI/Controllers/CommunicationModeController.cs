using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
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

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/communicationmode")]
    public class CommunicationModeController : ApiController
    {
        #region Create Communication Mode
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateCommunicationMode([FromBody]CommunicationModeModel communicationMode)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                if (!AccessControl.CanUserAccessPage(communicationMode.UserID, "Create Communication Mode"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCommunicationMode";

                    //Add parameters with values - Mode - 1 - Create
                    cmd.Parameters.AddWithValue("@CommunicationMode", communicationMode.CommunicationMode);
                    cmd.Parameters.AddWithValue("@IsActive", communicationMode.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", communicationMode.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Create);

                    //Calling sp to create entry
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

        #region Read Communication Modes
        [HttpGet]
        [Route("readcommunicationmodes/{UserID}/{isactiveonly?}")]
        public IHttpActionResult ReadCommunicationModes(string UserID, bool IsActiveOnly = false)
        {
            try
            {
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(UserID, "Communication Mode List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of communication modes
                List<CommunicationModeModel> CommunicationModeList = new List<CommunicationModeModel>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spCommunicationMode";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of communication modes
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        CommunicationModeModel communicationModeModel = new CommunicationModeModel();
                        communicationModeModel.SlNo = SlNo;
                        communicationModeModel.CommunicationModeID = Convert.ToInt32(sqlReader["CommunicationModeID"]);
                        communicationModeModel.CommunicationMode = sqlReader["CommunicationMode"].ToString();
                        communicationModeModel.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        CommunicationModeList.Add(communicationModeModel);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    if (IsActiveOnly)
                        return Ok(CommunicationModeList.Where(pa => pa.IsActive == true));
                    else
                        return Ok(CommunicationModeList);
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

        #region Read Communication Mode By Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult ReadCommunicationModeById(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Communication Mode"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a communication mode instance
                CommunicationModeModel communicationMode = new CommunicationModeModel();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spCommunicationMode";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@CommunicationModeID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get communication mode details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        communicationMode.CommunicationModeID = Convert.ToInt32(sqlReader["CommunicationModeID"]);
                        communicationMode.CommunicationMode = sqlReader["CommunicationMode"].ToString();
                        communicationMode.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        conn.Close();

                        //return communication mode to the request
                        return Ok(communicationMode);
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

        #region Update Communication Mode
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage UpdateCommunicationMode(int id, [FromBody]CommunicationModeModel communicationModeModel)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                //Check Communication Mode Id
                if (id != communicationModeModel.CommunicationModeID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Communication Mode id");

                if (!AccessControl.CanUserAccessPage(communicationModeModel.UserID, "Edit Communication Mode"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCommunicationMode";

                    //Add parameters with values - Mode - 2 - Update
                    cmd.Parameters.AddWithValue("@CommunicationModeID", communicationModeModel.CommunicationModeID);
                    cmd.Parameters.AddWithValue("@CommunicationMode", communicationModeModel.CommunicationMode);
                    cmd.Parameters.AddWithValue("@IsActive", communicationModeModel.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", communicationModeModel.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);

                    //Calling sp to update Communication Mode
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

        #region Delete Communication Mode
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteCommunicationMode(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete Communication Mode"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCommunicationMode";

                    //Add parameters with values - Mode - 3 - Delete
                    cmd.Parameters.AddWithValue("@CommunicationModeID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete Communication Mode
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.ToLower() == "deleted")
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

        #region Export To Excel
        [HttpGet]
        [Route("ExportCommunicationModeListToExcel")]
        public HttpResponseMessage ExportCommunicationModeListToExcel()
        {
            try
            {
                string FileName = "CommunicationModeList.xlsx";

                //Create a list to hold the list of Communication Modes
                List<CommunicationModeModel> communicationModeModeList = new List<CommunicationModeModel>();
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

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spCommunicationMode";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of Communication Modes
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            CommunicationModeModel communicationModeModel = new CommunicationModeModel();
                            communicationModeModel.CommunicationModeID = Convert.ToInt32(sqlReader["CommunicationModeID"]);
                            communicationModeModel.CommunicationMode = sqlReader["CommunicationMode"].ToString();
                            communicationModeModel.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                            communicationModeModeList.Add(communicationModeModel);
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Communication Mode ID");
                        ws.Cells[0, 2].PutValue("CommunicationMode");
                        ws.Cells[0, 3].PutValue("Is Active?");

                        for (int c = 0; c <= 3; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (CommunicationModeModel communicationModeModel in communicationModeModeList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(communicationModeModel.CommunicationModeID);
                            ws.Cells[row, 2].PutValue(communicationModeModel.CommunicationMode);
                            if (communicationModeModel.IsActive)
                                ws.Cells[row, 3].PutValue("Yes");
                            else
                                ws.Cells[row, 3].PutValue("No");
                            #endregion

                            #region setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 2].SetStyle(styleCenterAlignData);
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
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "No data found");
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
