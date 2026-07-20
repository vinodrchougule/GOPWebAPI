using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.UI.WebControls;
using ChangePassword = GOPWebAPI.Models.ChangePassword;
using Login = GOPWebAPI.Models.Login;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/login")]
    public class LoginController : ApiController
    {
        #region Validate Login Credentials
        [HttpPatch]
        [Route]
        public HttpResponseMessage ValidateLogin([FromBody] Login login)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Username or Password");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spValidateLogin";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@UserName", login.UserName);
                    cmd.Parameters.AddWithValue("@Password", AccountSecurity.EncodePasswordToBase64(login.Password));

                    //Calling sp to validate login
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
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

        #region Forgot Password
        [HttpPatch]
        [Route("forgotpassword/{username}")]
        public HttpResponseMessage ForgotPassword(string Username)
        {
            try
            {
                if (string.IsNullOrEmpty(Username) || Username.Trim().Length < 3 || Username.Length > 50)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Username");

                RandomGenerator randomGenerator = new RandomGenerator();
                string RandomPassword = randomGenerator.RandomPassword();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spForgotPassword";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@UserName", Username);
                    cmd.Parameters.AddWithValue("@Password", RandomPassword);
                    cmd.Parameters.AddWithValue("@EncodedPassword", AccountSecurity.EncodePasswordToBase64(RandomPassword));

                    //Calling sp to send password email
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
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

        #region Change Password
        [HttpPatch]
        [Route("changepassword")]
        public HttpResponseMessage ChangePassword(ChangePassword changePassword)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Username or Password");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spChangePassword";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@UserName", changePassword.UserName);
                    cmd.Parameters.AddWithValue("@Password", AccountSecurity.EncodePasswordToBase64(changePassword.Password));
                    cmd.Parameters.AddWithValue("@NewPassword", AccountSecurity.EncodePasswordToBase64(changePassword.NewPassword));

                    //Calling sp to change password
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
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

        #region Alert user about password expiry
        [HttpGet]
        [Route("AlertPasswordExpiry/{username}")]
        public HttpResponseMessage AlertPasswordExpiry(string Username)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spIsPasswordExpiring";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@Username", Username);

                    //Calling sp to check password expiry
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    //return response status code
                    string[] arrResult = Result.Split(',');
                    if (arrResult[0].Trim().ToLower() == "yes")
                    {
                        string ExpiryMessage = arrResult[1];
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ExpiryMessage);
                    }
                    else
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
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

        #region Reset Password
        [HttpPatch]
        [Route("ResetPassword/{Username}")]
        public HttpResponseMessage ResetPassword(string Username)
        {
            try
            {
                if (string.IsNullOrEmpty(Username) || Username.Trim().Length < 3 || Username.Length > 50)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Username");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spResetPassword";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@UserName", Username);

                    conn.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        if (Convert.ToBoolean(rdr["ReturnCode"]))
                        {
                            string UniqueId = rdr["UniqueId"].ToString();
                            conn.Close();

                            using (SqlConnection conn1 = new SqlConnection(DBConnInfo.ConnectionString()))
                            {
                                //send password to registered email id.
                                //Initialize command object
                                SqlCommand cmdSendPasswordLink = new SqlCommand();
                                cmdSendPasswordLink.Connection = conn1;
                                cmdSendPasswordLink.CommandType = CommandType.StoredProcedure;
                                cmdSendPasswordLink.CommandText = "spSendResetPasswordLink";

                                //Add parameters with values
                                cmdSendPasswordLink.Parameters.AddWithValue("@UserName", Username);
                                cmdSendPasswordLink.Parameters.AddWithValue("@UniqueId", UniqueId);

                                conn1.Open();
                                string Result = cmdSendPasswordLink.ExecuteScalar().ToString();
                                conn1.Close();

                                if (Result.ToLower() == "success")
                                    return Request.CreateResponse(HttpStatusCode.OK, Result);
                                else
                                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                            }
                        }
                        else
                        {
                            conn.Close();
                            string Result = "Username not found!";
                            return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
                        }
                    }

                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Request");
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

        #region Is Password Link Valid
        [HttpGet]
        [Route("IsPasswordResetLinkValid/{UId}")]
        public HttpResponseMessage IsPasswordResetLinkValid(string UId)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spIsPasswordResetLinkValid";

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@GUID", UId);

                    //Calling sp to validate login
                    conn.Open();
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if(Result.ToLower()=="success")
                        return Request.CreateResponse(HttpStatusCode.OK, Result);
                    else
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Password Reset link is invalid");
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

        #region Change Password Using Password Reset Link
        [HttpPatch]
        [Route("ChangePasswordUsingPasswordResetLink")]
        public HttpResponseMessage ChangePasswordUsingPasswordResetLink(ResetPassword resetPassword)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmdChangePasswordUsingResetLink = new SqlCommand();
                    cmdChangePasswordUsingResetLink.Connection = conn;
                    cmdChangePasswordUsingResetLink.CommandType = CommandType.StoredProcedure;
                    cmdChangePasswordUsingResetLink.CommandText = "spChangePasswordUsingResetLink";

                    //Add parameters with values
                    cmdChangePasswordUsingResetLink.Parameters.AddWithValue("@GUID", resetPassword.UId);
                    cmdChangePasswordUsingResetLink.Parameters.AddWithValue("@Password", AccountSecurity.EncodePasswordToBase64(resetPassword.Password));

                    //Calling sp to change password
                    conn.Open();
                    string Result = cmdChangePasswordUsingResetLink.ExecuteScalar().ToString();
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
