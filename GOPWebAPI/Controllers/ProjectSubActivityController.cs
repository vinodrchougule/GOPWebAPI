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
    [RoutePrefix("api/projectsubactivity")]
    public class ProjectSubActivityController : ApiController
    {
        #region Create Project Sub-Activity
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateProjectSubActivity([FromBody]ProjectSubActivity projectSubActivity)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                if (!AccessControl.CanUserAccessPage(projectSubActivity.UserID, "Create Project SubActivity"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectSubActivity";

                    //Add parameters with values - Mode - 1 - Create
                    cmd.Parameters.AddWithValue("@SubActivity", projectSubActivity.SubActivity);
                    cmd.Parameters.AddWithValue("@IsActive", projectSubActivity.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", projectSubActivity.UserID);
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

        #region Read Project Sub-Activities
        [HttpGet]
        [Route("{UserID}")]
        public IHttpActionResult ReadProjectSubActivities(string UserID)
        {
            try
            {
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(UserID, "Project SubActivity List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of project sub activities
                List<ProjectSubActivity> ProjectSubActivityList = new List<ProjectSubActivity>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectSubActivity";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of project sub-activities
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectSubActivity projectSubActivity = new ProjectSubActivity();
                        projectSubActivity.SlNo = SlNo;
                        projectSubActivity.ProjectSubActivityID = Convert.ToInt32(sqlReader["ProjectSubActivityID"]);
                        projectSubActivity.SubActivity = sqlReader["SubActivity"].ToString();
                        projectSubActivity.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        ProjectSubActivityList.Add(projectSubActivity);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ProjectSubActivityList);
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

        #region Read Project Sub-Activity By Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult ReadProjectSubActivityById(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Project SubActivity"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a project sub activity instance
                ProjectSubActivity projectSubActivity = new ProjectSubActivity();
                System.Data.Common.DbDataReader sqlReader;
                
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectSubActivity";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@ProjectSubActivityID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get project sub-activity details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        projectSubActivity.ProjectSubActivityID = Convert.ToInt32(sqlReader["ProjectSubActivityID"]);
                        projectSubActivity.SubActivity = sqlReader["SubActivity"].ToString();
                        projectSubActivity.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        conn.Close();

                        //return project sub-activity to the request
                        return Ok(projectSubActivity);
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

        #region Update Project Sub-Activity
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage UpdateProjectSubActivity(int id, [FromBody]ProjectSubActivity projectSubActivity)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                //Check Project Sub-Activity Id
                if (id != projectSubActivity.ProjectSubActivityID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Project Sub-Activity id");

                if (!AccessControl.CanUserAccessPage(projectSubActivity.UserID, "Edit Project SubActivity"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectSubActivity";

                    //Add parameters with values - Mode - 2 - Update
                    cmd.Parameters.AddWithValue("@ProjectSubActivityID", projectSubActivity.ProjectSubActivityID);
                    cmd.Parameters.AddWithValue("@SubActivity", projectSubActivity.SubActivity);
                    cmd.Parameters.AddWithValue("@IsActive", projectSubActivity.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", projectSubActivity.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);

                    //Calling sp to update project sub-activity
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

        #region Delete Project Sub-Activity
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteProjectSubActivity(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete Project SubActivity"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spProjectSubActivity";

                    //Add parameters with values - Mode - 3 - Delete
                    cmd.Parameters.AddWithValue("@ProjectSubActivityID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete Project Sub-Activity
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
        [Route("ExportProjectSubActivityListToExcel")]
        public HttpResponseMessage ExportProjectSubActivityListToExcel()
        {
            try
            {
                string FileName = "ProjectSubActivitiesList.xlsx";

                //Create a list to hold the list of Project Sub-Activities
                List<ProjectSubActivity> projectSubActivityList = new List<ProjectSubActivity>();
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
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectSubActivity";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of Project Sub-Activities
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            ProjectSubActivity projectSubActivity = new ProjectSubActivity();
                            projectSubActivity.ProjectSubActivityID = Convert.ToInt32(sqlReader["ProjectSubActivityID"]);
                            projectSubActivity.SubActivity = sqlReader["SubActivity"].ToString();
                            projectSubActivity.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                            projectSubActivityList.Add(projectSubActivity);
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Project Sub-Activity ID");
                        ws.Cells[0, 2].PutValue("Sub-Activity");
                        ws.Cells[0, 3].PutValue("Is Active?");

                        for (int c = 0; c <= 3; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (ProjectSubActivity projectSubActivity in projectSubActivityList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(projectSubActivity.ProjectSubActivityID);
                            ws.Cells[row, 2].PutValue(projectSubActivity.SubActivity);
                            if (projectSubActivity.IsActive)
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
