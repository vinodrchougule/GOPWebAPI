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
    [RoutePrefix("api/RoleAccess")]
    public class RoleAccessController : ApiController
    {
        #region Update Role Access
        [HttpPost]
        [Route]
        public HttpResponseMessage UpdateRoleAccess([FromBody]RoleAccess roleAccess)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Role Access Data");

                if (!AccessControl.CanUserAccessPage(roleAccess.UserID, "Create-Edit Role Access"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Role Access List
                //create xml serialized data
                var serializer = new XmlSerializer(typeof(List<RolePageAccess>), new XmlRootAttribute("root"));

                //create a stream
                var stream = new StringWriter();

                //write serialized data to stream
                serializer.Serialize(stream, roleAccess.RoleAccessList);

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
                    cmd.CommandText = "spRoleAccess";

                    #region Adding Stored Procedure Parameters
                    //Send Role Access Details as xml
                    cmd.Parameters.Add(new SqlParameter("@RoleAccessDetails", SqlDbType.Xml)
                    {
                        Value = sqlXml
                    });

                    cmd.Parameters.AddWithValue("@UserID", roleAccess.UserID);
                    #endregion

                    //Calling sp to update role access
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

        #region Read Role Access List
        [HttpGet]
        [Route("RoleAccessList/{UserID}")]
        public IHttpActionResult ReadRoleAccessList(string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Role Access List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of Role Access
                List<RolePageAccess> RoleAccessList = new List<RolePageAccess>();
                System.Data.Common.DbDataReader sqlReader;
                int SlNo = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spRoleAccessList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of Role Access
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        RolePageAccess rolePageAccess = new RolePageAccess();

                        //Assign values to model object
                        rolePageAccess.SlNo = SlNo;
                        rolePageAccess.RoleName = sqlReader["RoleName"].ToString();
                        rolePageAccess.PageName = sqlReader["PageName"].ToString();
                        rolePageAccess.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);

                        //Add object to list
                        RoleAccessList.Add(rolePageAccess);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(RoleAccessList);
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

        #region Read Role Access By RoleName
        [HttpGet]
        [Route("{RoleName}/{UserID}")]
        public IHttpActionResult ReadRoleAccessByRoleName(string RoleName, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Role Access"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of Role Access
                List<RolePageAccess> RoleAccessList = new List<RolePageAccess>();
                System.Data.Common.DbDataReader sqlReader;
                int SlNo = 1;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spRoleAccessByRoleName";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@RoleName", RoleName);
                    #endregion

                    //Calling sp to get list of Role Access
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        //Create model object
                        RolePageAccess rolePageAccess = new RolePageAccess();

                        //Assign values to model object
                        rolePageAccess.SlNo = SlNo;
                        rolePageAccess.RoleName = sqlReader["RoleName"].ToString();
                        rolePageAccess.PageName = sqlReader["PageName"].ToString();
                        rolePageAccess.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);

                        //Add object to list
                        RoleAccessList.Add(rolePageAccess);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(RoleAccessList);
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
        [Route("ExportRoleAccessListToExcel")]
        public HttpResponseMessage ExportRoleAccessListToExcel()
        {
            try
            {
                int SlNo = 1, row = 1;
                string FileName = "RoleAccessList.xlsx";

                //Create a list to hold the list of Role Access
                List<RolePageAccess> RoleAccessList = new List<RolePageAccess>();
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
                    cmd.CommandText = "spRoleAccessList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of Role Access
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            //Create model object
                            RolePageAccess rolePageAccess = new RolePageAccess();

                            //Assign values to model object
                            rolePageAccess.SlNo = SlNo;
                            rolePageAccess.RoleName = sqlReader["RoleName"].ToString();
                            rolePageAccess.PageName = sqlReader["PageName"].ToString();
                            rolePageAccess.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);

                            //Add object to list
                            RoleAccessList.Add(rolePageAccess);
                            SlNo++;
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Role Name");
                        ws.Cells[0, 2].PutValue("Page Name");
                        ws.Cells[0, 3].PutValue("Is Active?");

                        for (int c = 0; c <= 3; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (RolePageAccess rolePageAccess in RoleAccessList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(rolePageAccess.SlNo);
                            ws.Cells[row, 1].PutValue(rolePageAccess.RoleName);
                            ws.Cells[row, 2].PutValue(rolePageAccess.PageName);
                            if (rolePageAccess.IsActive)
                                ws.Cells[row, 3].PutValue("Yes");
                            else
                                ws.Cells[row, 3].PutValue("No");
                            #endregion

                            #region Setting row data style
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
