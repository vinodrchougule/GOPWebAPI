using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web;
using System.Xml;
using System.Xml.Serialization;
using System.Data.SqlTypes;
using System.Net.Http.Headers;
using GOPWebAPI.Helpers;
using Aspose.Cells;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/UserRole")]
    public class UserRoleController : ApiController
    {
        #region Update User Role(s)
        [HttpPost]
        [Route]
        public HttpResponseMessage UpdateUserRole([FromBody]UserRoles userRoles)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid User Roles Data");

                if (!AccessControl.CanUserAccessPage(userRoles.UserID, "Create-Edit User Role(s)"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region User Role List
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<UserRole>),
                                       new XmlRootAttribute("root"));

                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, userRoles.UserRolesList);

                //Read stream as xml string
                StringReader transactionXml = new StringReader(stream.ToString());

                //Read xml string
                XmlTextReader xmlReader = new XmlTextReader(transactionXml);

                //Convert xml string to sql xml
                SqlXml sqlXml = new SqlXml(xmlReader);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUserRole";

                    #region Adding Stored Procedure Parameters
                    //Send User Role Details as xml
                    cmd.Parameters.Add(new SqlParameter("@UserRoleDetails", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });

                    cmd.Parameters.AddWithValue("@UserID", userRoles.UserID);
                    #endregion

                    //Calling sp to update user role(s)
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.Trim().ToLower() == "updated")
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

        #region Read User Role List
        [HttpGet]
        [Route("UserRoleList/{UserID}")]
        public IHttpActionResult ReadUserRoleList(string UserID)
        {
            try
            {
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(UserID, "User Role(s) List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of User Roles
                List<UserRole> UserRoleList = new List<UserRole>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUserRoleList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of User Roles
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        UserRole userRole = new UserRole();

                        //Assign values to model object
                        userRole.SlNo = SlNo;
                        userRole.UserName = sqlReader["UserName"].ToString();
                        userRole.RoleName = sqlReader["RoleName"].ToString();
                        userRole.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);

                        //Add object to list
                        UserRoleList.Add(userRole);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(UserRoleList);
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

        #region Read User Roles By UserName
        [HttpGet]
        [Route("{UserName}/{UserID}")]
        public IHttpActionResult ReadUserRolesByUserName(string UserName, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View User Role(s)"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of User Roles
                List<UserRole> UserRoles = new List<UserRole>();
                System.Data.Common.DbDataReader sqlReader;
                int SlNo = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUserRolesByUsername";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@UserName", UserName);
                    #endregion

                    //Calling sp to get list of User Roles
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        UserRole userRole = new UserRole();

                        //Assign values to model object
                        userRole.SlNo = SlNo;
                        userRole.UserName = sqlReader["UserName"].ToString();
                        userRole.RoleName = sqlReader["RoleName"].ToString();
                        userRole.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);

                        //Add object to list
                        UserRoles.Add(userRole);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(UserRoles);
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

        #region User Roles Pagewise Access List
        [HttpGet]
        [Route("ReadUserRolesPagewiseAccess/{UserName}")]
        public IHttpActionResult ReadUserRolesPagewiseAccess(string UserName)
        {
            try
            {
                //Create a list to hold the list of User Roles Pagewise Access
                List<UserRolePageAccess> UserRolePageAccessList = new List<UserRolePageAccess>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUserRolesPagewiseAccessList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@UserName", UserName);
                    #endregion

                    //Calling sp to get list of User Roles
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        UserRolePageAccess userRolePageAccess = new UserRolePageAccess();

                        //Assign values to model object
                        userRolePageAccess.RoleName = sqlReader["RoleName"].ToString();
                        userRolePageAccess.PageName = sqlReader["PageName"].ToString();
                        userRolePageAccess.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);

                        //Add object to list
                        UserRolePageAccessList.Add(userRolePageAccess);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(UserRolePageAccessList);
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

        #region Export To Excel
        [HttpGet]
        [Route("ExportUserRolesListToExcel")]
        public HttpResponseMessage ExportUserRolesListToExcel()
        {
            try
            {
                int SlNo = 1, row = 1;
                string FileName = "User Roles List.xlsx";

                //Create a list to hold the list of User Roles
                List<UserRole> UserRoleList = new List<UserRole>();
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
                    cmd.CommandText = "spUserRoleList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of User Roles
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            //Create model object
                            UserRole userRole = new UserRole();

                            //Assign values to model object
                            userRole.SlNo = SlNo;
                            userRole.UserName = sqlReader["UserName"].ToString();
                            userRole.RoleName = sqlReader["RoleName"].ToString();
                            userRole.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);

                            //Add object to list
                            UserRoleList.Add(userRole);
                            SlNo++;
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("User Name");
                        ws.Cells[0, 2].PutValue("Role Name");
                        ws.Cells[0, 3].PutValue("Is Active?");

                        for (int c = 0; c <= 3; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (UserRole userRole in UserRoleList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(userRole.SlNo);
                            ws.Cells[row, 1].PutValue(userRole.UserName);
                            ws.Cells[row, 2].PutValue(userRole.RoleName);
                            if (userRole.IsActive)
                                ws.Cells[row, 3].PutValue("Yes");
                            else
                                ws.Cells[row, 3].PutValue("No");
                            #endregion

                            #region Setting row data style
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
