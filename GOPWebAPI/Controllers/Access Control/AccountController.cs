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
using System.Security.Cryptography;
using System.Security.Principal;
using System.Web;
using System.Web.Http;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/account")]
    public class AccountController : ApiController
    {
        #region Create Account
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateAccount([FromBody]Account account)
        {
            try
            {
                string PhotoFileExtension = string.Empty;

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                if (!AccessControl.CanUserAccessPage(account.User, "Create User"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                //Photo File Extension
                if (File.Exists(dirTemp + account.PhotoFileName))
                    PhotoFileExtension = Path.GetExtension(dirTemp + account.PhotoFileName);

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUser";

                    //Add parameters with values - Mode - 1 - Create
                    cmd.Parameters.AddWithValue("@FirstName", account.FirstName);
                    cmd.Parameters.AddWithValue("@MiddleName", account.MiddleName);
                    cmd.Parameters.AddWithValue("@LastName", account.LastName);
                    cmd.Parameters.AddWithValue("@UserName", account.UserName);
                    cmd.Parameters.AddWithValue("@Password", AccountSecurity.EncodePasswordToBase64(account.Password));
                    cmd.Parameters.AddWithValue("@Email", account.Email);
                    cmd.Parameters.AddWithValue("@DepartmentName", account.DepartmentName);
                    cmd.Parameters.AddWithValue("@ManagerName", account.ManagerName);
                    cmd.Parameters.AddWithValue("@PhotoFileNameExtension", PhotoFileExtension);
                    cmd.Parameters.AddWithValue("@User", account.User);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Create);

                    //Calling sp to create account
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.ToLower() == "created")
                    {
                        //----------------------------------Photo File move Starts------------------------------------------------//
                        if (!string.IsNullOrEmpty(account.PhotoFileName))
                        {
                            #region Photo File move Starts
                            if (File.Exists(dirTemp + account.PhotoFileName))
                            {
                                string NewPhotoFileName = account.UserName.ToUpper() + PhotoFileExtension;
                                DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/UserImages/"));
                                FileOperations.MoveFile(dirTemp, account.PhotoFileName, dirUploads, NewPhotoFileName);
                            }
                            #endregion
                        }
                        //----------------------------------Photo File move Ends---------------------------------------------------//

                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
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

        #region Read Accounts
        [HttpGet]
        [Route("{UserID}")]
        public IHttpActionResult ReadAccounts(string UserID)
        {
            try
            {
                int SlNo = 1;
                string photoFileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/UserImages/");

                if (!AccessControl.CanUserAccessPage(UserID, "User List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of accounts
                List<Account> AccountList = new List<Account>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUser";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of accounts
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        Account account = new Account();
                        account.SlNo = SlNo;
                        account.UserID = Convert.ToInt32(sqlReader["UserID"]);
                        account.FirstName = sqlReader["FirstName"].ToString();
                        account.MiddleName = sqlReader["MiddleName"].ToString();
                        account.LastName = sqlReader["LastName"].ToString();
                        account.UserName = sqlReader["UserName"].ToString();
                        account.IsLockedOut = Convert.ToBoolean(sqlReader["IsLockedOut"]);
                        account.Email = sqlReader["Email"].ToString();
                        account.DepartmentName = sqlReader["Department"].ToString();
                        account.ManagerName = sqlReader["ManagerName"].ToString();
                        account.RelievingDate = sqlReader["RelievedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["RelievedOn"];
                        account.PhotoFileName = sqlReader["PhotoFileName"].ToString();
                        AccountList.Add(account);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(AccountList);
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

        #region Read Account By Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult AccountById(int id, string UserID)
        {
            try
            {
                string photoFileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/UserImages/");

                if (!AccessControl.CanUserAccessPage(UserID, "View User"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a account instance
                Account account = new Account();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUser";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@UserID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get account details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        account.UserID = Convert.ToInt32(sqlReader["UserID"]);
                        account.FirstName = sqlReader["FirstName"].ToString();
                        account.MiddleName = sqlReader["MiddleName"].ToString();
                        account.LastName = sqlReader["LastName"].ToString();
                        account.UserName = sqlReader["UserName"].ToString();
                        account.IsLockedOut = Convert.ToBoolean(sqlReader["IsLockedOut"]);
                        account.Email = sqlReader["Email"].ToString();
                        account.DepartmentName = sqlReader["Department"].ToString();
                        account.ManagerName = sqlReader["ManagerName"].ToString();
                        account.RelievingDate = sqlReader["RelievedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["RelievedOn"];
                        account.PhotoFileName = sqlReader["PhotoFileName"].ToString();
                        if (!string.IsNullOrEmpty(account.PhotoFileName) && File.Exists(photoFileUploadedPath + account.PhotoFileName))
                        {
                            string photoFilePath = Path.Combine(photoFileUploadedPath, account.PhotoFileName);
                            FileInfo fileInfo = new FileInfo(photoFilePath);
                            //Read the Image as Byte Array.
                            byte[] bytes = File.ReadAllBytes(fileInfo.FullName);
                            account.PhotoFileBase64String = Convert.ToBase64String(bytes, 0, bytes.Length);
                        }
                        conn.Close();

                        //return account to the request
                        return Ok(account);
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

        #region Read Account Details By Username
        [HttpGet]
        [Route("accountbyusername/{Username}")]
        public IHttpActionResult AccountByUsername(string Username)
        {
            try
            {
                string photoFileUploadedPath = HttpContext.Current.Server.MapPath("~/Uploads/UserImages/");

                //Create a account instance
                Account account = new Account();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUserDetailsByUsername";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@Username", Username);

                    //Calling sp to get account details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        account.UserID = Convert.ToInt32(sqlReader["UserID"]);
                        account.FirstName = sqlReader["FirstName"].ToString();
                        account.MiddleName = sqlReader["MiddleName"].ToString();
                        account.LastName = sqlReader["LastName"].ToString();
                        account.UserName = sqlReader["UserName"].ToString();
                        account.IsLockedOut = Convert.ToBoolean(sqlReader["IsLockedOut"]);
                        account.Email = sqlReader["Email"].ToString();
                        account.DepartmentName = sqlReader["Department"].ToString();
                        account.ManagerName = sqlReader["ManagerName"].ToString();
                        account.RelievingDate = sqlReader["RelievedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["RelievedOn"];
                        account.PhotoFileName = sqlReader["PhotoFileName"].ToString();
                        if (!string.IsNullOrEmpty(account.PhotoFileName) && File.Exists(photoFileUploadedPath + account.PhotoFileName))
                        {
                            string photoFilePath = Path.Combine(photoFileUploadedPath, account.PhotoFileName);
                            FileInfo fileInfo = new FileInfo(photoFilePath);
                            //Read the Image as Byte Array.
                            byte[] bytes = File.ReadAllBytes(fileInfo.FullName);
                            account.PhotoFileBase64String = Convert.ToBase64String(bytes, 0, bytes.Length);
                        }
                        conn.Close();

                        //return account to the request
                        return Ok(account);
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

        #region Read Accounts List By Department
        [HttpGet]
        [Route("AccountsListByDepartment")]
        public IHttpActionResult AccountsListByDepartment(string Department)
        {
            try
            {
                List<string> accountsList = new List<string>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUsersListByDepartment";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@Department", Department);

                    //Calling sp to get account details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        accountsList.Add(sqlReader["UserName"].ToString());
                    }
                    conn.Close();

                    //return account to the request
                    return Ok(accountsList);
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

        #region Update Account
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage UpdateAccount(int id, [FromBody]Account account)
        {
            try
            {
                string PhotoFileExtension = string.Empty;

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                //Check account Id
                if (id != account.UserID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Account id");

                if (!AccessControl.CanUserAccessPage(account.User, "Edit User"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));
                DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/UserImages/"));

                //Photo File Extension
                if (File.Exists(dirTemp + account.PhotoFileName))
                    PhotoFileExtension = Path.GetExtension(dirTemp + account.PhotoFileName).ToLower();
                else if(!string.IsNullOrEmpty(account.PhotoFileName))
                    PhotoFileExtension = account.PhotoFileName.Substring(account.PhotoFileName.IndexOf('.'));

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUser";

                    //Add parameters with values - Mode - 2 - Update
                    cmd.Parameters.AddWithValue("@UserID", account.UserID);
                    cmd.Parameters.AddWithValue("@FirstName", account.FirstName);
                    cmd.Parameters.AddWithValue("@MiddleName", account.MiddleName);
                    cmd.Parameters.AddWithValue("@LastName", account.LastName);
                    cmd.Parameters.AddWithValue("@UserName", account.UserName);
                    cmd.Parameters.AddWithValue("@Email", account.Email);
                    cmd.Parameters.AddWithValue("@DepartmentName", account.DepartmentName);
                    cmd.Parameters.AddWithValue("@ManagerName", account.ManagerName);
                    cmd.Parameters.AddWithValue("@PhotoFileNameExtension", PhotoFileExtension);
                    cmd.Parameters.AddWithValue("@RelievedOn", account.RelievingDate);
                    cmd.Parameters.AddWithValue("@User", account.User);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);

                    //Calling sp to update account
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.ToLower() == "updated")
                    {
                        //----------------------------------Photo File move Starts------------------------------------------------//
                        if (!string.IsNullOrEmpty(account.PhotoFileName))
                        {
                            #region Photo File move Starts
                            if (File.Exists(dirTemp + account.PhotoFileName))
                            {
                                string NewPhotoFileName = account.UserName.ToUpper() + PhotoFileExtension;
                                FileOperations.MoveFile(dirTemp, account.PhotoFileName, dirUploads, NewPhotoFileName);
                            }
                            #endregion
                        }
                        //----------------------------------Photo File move Ends---------------------------------------------------//

                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    }
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

        #region Delete Account
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteAccount(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete User"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spUser";

                    //Add parameters with values - Mode - 3 - Delete
                    cmd.Parameters.AddWithValue("@UserID", id);
                    cmd.Parameters.AddWithValue("@User", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete user
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
        [Route("ExportUsersListToExcel/{DepartmentName}/{ManagerName}")]
        public HttpResponseMessage ExportUsersListToExcel(string DepartmentName, string ManagerName)
        {
            try
            {
                int SlNo = 1, row = 1;
                string FileName = "UsersList.xlsx";

                //Create a list to hold the list of accounts
                List<Account> AccountList = new List<Account>();
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
                    cmd.CommandText = "spUser";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - DepartmentName, ManagerName, Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@DepartmentName", DepartmentName);
                    cmd.Parameters.AddWithValue("@ManagerName", ManagerName);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of accounts
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            Account account = new Account();
                            account.SlNo = SlNo;
                            account.UserID = Convert.ToInt32(sqlReader["UserID"]);
                            account.FirstName = sqlReader["FirstName"].ToString();
                            account.MiddleName = sqlReader["MiddleName"].ToString();
                            account.LastName = sqlReader["LastName"].ToString();
                            account.UserName = sqlReader["UserName"].ToString();
                            account.IsLockedOut = Convert.ToBoolean(sqlReader["IsLockedOut"]);
                            account.Email = sqlReader["Email"].ToString();
                            account.DepartmentName = sqlReader["Department"].ToString();
                            account.ManagerName = sqlReader["ManagerName"].ToString();
                            account.RelievingDate = sqlReader["RelievedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["RelievedOn"];
                            AccountList.Add(account);
                            SlNo++;
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("User ID");
                        ws.Cells[0, 2].PutValue("Full Name");
                        ws.Cells[0, 3].PutValue("User Name");
                        ws.Cells[0, 4].PutValue("Is Locked Out?");
                        ws.Cells[0, 5].PutValue("Email");
                        ws.Cells[0, 6].PutValue("Department");
                        ws.Cells[0, 7].PutValue("ManagerName");
                        ws.Cells[0, 8].PutValue("Relieving Date");

                        for (int c = 0; c <= 8; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (Account account in AccountList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(account.SlNo);
                            ws.Cells[row, 1].PutValue(account.UserID);
                            ws.Cells[row, 2].PutValue(account.FirstName + " " + account.MiddleName + " " + account.LastName);
                            ws.Cells[row, 3].PutValue(account.UserName);
                            if (account.IsLockedOut)
                                ws.Cells[row, 4].PutValue("Yes");
                            else
                                ws.Cells[row, 4].PutValue("No");
                            ws.Cells[row, 5].PutValue(account.Email);
                            ws.Cells[row, 6].PutValue(account.DepartmentName);
                            ws.Cells[row, 7].PutValue(account.ManagerName);
                            if (account.RelievingDate != null)
                                ws.Cells[row, 8].PutValue(Convert.ToDateTime(account.RelievingDate).ToString("dd-MMM-yyyy"));
                            #endregion

                            #region setting row data style
                            ws.Cells[row, 0].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 1].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 2].SetStyle(styleLeftAlignData);
                            ws.Cells[row, 3].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 4].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 5].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 6].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 7].SetStyle(styleCenterAlignData);
                            ws.Cells[row, 8].SetStyle(styleCenterAlignData);
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

        #region Read Accounts By Page No. (for Lazy Loading - not used it yet in LIVE project)
        [HttpGet]
        [Route("ReadAccountsByPageNo/{PageNo}/{PageSize}")]
        public IHttpActionResult ReadAccountsByPageNo(int PageNo, int PageSize)
        {
            try
            {
                //Create a list to hold the list of accounts
                List<Account> AccountList = new List<Account>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUserDetailsByPageNo";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@PageNo", PageNo);
                    cmd.Parameters.AddWithValue("@PageSize", PageSize);
                    //Calling sp to get list of accounts
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        Account account = new Account();
                        account.UserID = Convert.ToInt32(sqlReader["UserID"]);
                        account.FirstName = sqlReader["FirstName"].ToString();
                        account.MiddleName = sqlReader["MiddleName"].ToString();
                        account.LastName = sqlReader["LastName"].ToString();
                        account.UserName = sqlReader["UserName"].ToString();
                        account.IsLockedOut = Convert.ToBoolean(sqlReader["IsLockedOut"]);
                        account.Email = sqlReader["Email"].ToString();
                        AccountList.Add(account);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(AccountList);
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

        #region Reset User Credentials
        [HttpPatch]
        [Route("ResetUserCredentials")]
        public HttpResponseMessage ResetUserCredentials(int id, string UserName) //id - Id of user to reset the credentials, UserName - Login User
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmdResetUserCredentials = new SqlCommand();
                    cmdResetUserCredentials.Connection = conn;
                    cmdResetUserCredentials.CommandType = CommandType.StoredProcedure;
                    cmdResetUserCredentials.CommandText = "spResetUserCredentials";

                    //Add parameters with values
                    cmdResetUserCredentials.Parameters.AddWithValue("@UserID", id);
                    cmdResetUserCredentials.Parameters.AddWithValue("@UserName", UserName);

                    //Calling sp to reset user credentials
                    conn.Open();
                    string Result = cmdResetUserCredentials.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    if (Result.ToLower() == "success")
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
    }
}
