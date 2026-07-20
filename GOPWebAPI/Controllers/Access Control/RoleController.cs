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
    [RoutePrefix("api/role")]
    public class RoleController : ApiController
    {
        #region Create Role
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateRole([FromBody]Role role)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                if (!AccessControl.CanUserAccessPage(role.UserID, "Create Role"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spRole";

                    //Add parameters with values - Mode - 1 - Create
                    cmd.Parameters.AddWithValue("@RoleName", role.RoleName);
                    cmd.Parameters.AddWithValue("@IsActive", role.IsActive);
                    cmd.Parameters.AddWithValue("@Description", role.Description);
                    cmd.Parameters.AddWithValue("@UserID", role.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Create);

                    //Calling sp to create role
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

        #region Read Roles
        [HttpGet]
        [Route("roles/{UserID}/{isactiveonly?}")]
        public IHttpActionResult Roles(string UserID, bool IsActiveOnly = false)
        {
            try
            {
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(UserID, "Role List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of roles
                List<Role> RoleList = new List<Role>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spRole";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of roles
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        Role role = new Role();
                        role.SlNo = SlNo;
                        role.RoleID = Convert.ToInt32(sqlReader["RoleID"]);
                        role.RoleName = sqlReader["RoleName"].ToString();
                        role.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        role.Description = sqlReader["Description"].ToString();
                        RoleList.Add(role);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    if (IsActiveOnly)
                        return Ok(RoleList.Where(pa => pa.IsActive == true));
                    else
                        return Ok(RoleList);
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

        #region Read Role By Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult RoleById(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Role"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a role instance
                Role role = new Role();
                System.Data.Common.DbDataReader sqlReader;
                
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spRole";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@RoleID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get role details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        role.RoleID = Convert.ToInt32(sqlReader["RoleID"]);
                        role.RoleName = sqlReader["RoleName"].ToString();
                        role.Description = sqlReader["Description"].ToString();
                        role.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        conn.Close();

                        //return role to the request
                        return Ok(role);
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

        #region Update Role
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage PutRole(int id, [FromBody]Role role)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                //Check Role Id
                if (id != role.RoleID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Role id");

                if (!AccessControl.CanUserAccessPage(role.UserID, "Edit Role"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spRole";

                    //Add parameters with values - Mode - 2 - Update
                    cmd.Parameters.AddWithValue("@RoleID", role.RoleID);
                    cmd.Parameters.AddWithValue("@RoleName", role.RoleName);
                    cmd.Parameters.AddWithValue("@IsActive", role.IsActive);
                    cmd.Parameters.AddWithValue("@Description", role.Description);
                    cmd.Parameters.AddWithValue("@UserID", role.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);

                    //Calling sp to update role
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

        #region Delete Role
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteRole(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete Role"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spRole";

                    //Add parameters with values - Mode - 3 - Delete
                    cmd.Parameters.AddWithValue("@RoleID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete role
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
        [Route("ExportRolesListToExcel")]
        public HttpResponseMessage ExportRolesListToExcel()
        {
            try
            {
                int SlNo = 1, row = 1;
                string FileName = "RolesList.xlsx";

                //Create a list to hold the list of roles
                List<Role> RoleList = new List<Role>();
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
                    cmd.CommandText = "spRole";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of roles
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            Role role = new Role();
                            role.SlNo = SlNo;
                            role.RoleID = Convert.ToInt32(sqlReader["RoleID"]);
                            role.RoleName = sqlReader["RoleName"].ToString();
                            role.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                            role.Description = sqlReader["Description"].ToString();
                            RoleList.Add(role);
                            SlNo++;
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Role ID");
                        ws.Cells[0, 2].PutValue("Role Name");
                        ws.Cells[0, 3].PutValue("Description");
                        ws.Cells[0, 4].PutValue("Is Active?");
                        
                        for (int c = 0; c <= 4; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (Role role in RoleList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(role.SlNo);
                            ws.Cells[row, 1].PutValue(role.RoleID);
                            ws.Cells[row, 2].PutValue(role.RoleName);
                            ws.Cells[row, 3].PutValue(role.Description);
                            if (role.IsActive)
                                ws.Cells[row, 4].PutValue("Yes");
                            else
                                ws.Cells[row, 4].PutValue("No");
                            #endregion

                            #region Setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 2].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 3].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 4].SetStyle(styleCenterAlignData);
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
