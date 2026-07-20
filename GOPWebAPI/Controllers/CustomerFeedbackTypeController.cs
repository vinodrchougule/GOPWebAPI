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
    [RoutePrefix("api/customerfeedbacktype")]
    public class CustomerFeedbackTypeController : ApiController
    {
        #region Create Customer Feedback Type
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateCustomerFeedbackType([FromBody]CustomerFeedbackType customerFeedbackType)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                if (!AccessControl.CanUserAccessPage(customerFeedbackType.UserID, "Create Customer Feedback Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCustomerFeedbackType";

                    //Add parameters with values - Mode - 1 - Create
                    cmd.Parameters.AddWithValue("@FeedbackType", customerFeedbackType.FeedbackType);
                    cmd.Parameters.AddWithValue("@IsActive", customerFeedbackType.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", customerFeedbackType.UserID);
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

        #region Read Customer Feedback Types
        [HttpGet]
        [Route("readcustomerfeedbacktypes/{UserID}/{isactiveonly?}")]
        public IHttpActionResult ReadCustomerFeedbackTypes(string UserID, bool IsActiveOnly = false)
        {
            try
            {
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(UserID, "Customer Feedback Type List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of Customer Feedback Types
                List<CustomerFeedbackType> CustomerFeedbackTypesList = new List<CustomerFeedbackType>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spCustomerFeedbackType";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of Customer Feedback Types
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        CustomerFeedbackType customerFeedbackType = new CustomerFeedbackType();
                        customerFeedbackType.SlNo = SlNo;
                        customerFeedbackType.CustomerFeedbackTypeID = Convert.ToInt32(sqlReader["CustomerFeedbackTypeID"]);
                        customerFeedbackType.FeedbackType = sqlReader["FeedbackType"].ToString();
                        customerFeedbackType.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        CustomerFeedbackTypesList.Add(customerFeedbackType);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    if (IsActiveOnly)
                        return Ok(CustomerFeedbackTypesList.Where(pa => pa.IsActive == true));
                    else
                        return Ok(CustomerFeedbackTypesList);
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

        #region Read Customer Feedback Type By Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult ReadCustomerFeedbackTypeById(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Customer Feedback Type"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a Customer Feedback Type instance
                CustomerFeedbackType customerFeedbackType = new CustomerFeedbackType();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spCustomerFeedbackType";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@CustomerFeedbackTypeID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get Customer Feedback Type details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        customerFeedbackType.CustomerFeedbackTypeID = Convert.ToInt32(sqlReader["CustomerFeedbackTypeID"]);
                        customerFeedbackType.FeedbackType = sqlReader["FeedbackType"].ToString();
                        customerFeedbackType.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        conn.Close();

                        //return customer feedback type to the request
                        return Ok(customerFeedbackType);
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

        #region Update Customer Feedback Type
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage UpdateCustomerFeedbackType(int id, [FromBody]CustomerFeedbackType customerFeedbackType)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                //Check Project Activity Id
                if (id != customerFeedbackType.CustomerFeedbackTypeID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Customer Feedback Type id");

                if (!AccessControl.CanUserAccessPage(customerFeedbackType.UserID, "Edit Customer Feedback Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCustomerFeedbackType";

                    //Add parameters with values - Mode - 2 - Update
                    cmd.Parameters.AddWithValue("@CustomerFeedbackTypeID", customerFeedbackType.CustomerFeedbackTypeID);
                    cmd.Parameters.AddWithValue("@FeedbackType", customerFeedbackType.FeedbackType);
                    cmd.Parameters.AddWithValue("@IsActive", customerFeedbackType.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", customerFeedbackType.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);

                    //Calling sp to update customer feedback type
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

        #region Delete Customer Feedback Type
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteCustomerFeedbackType(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete Customer Feedback Type"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spCustomerFeedbackType";

                    //Add parameters with values - Mode - 3 - Delete
                    cmd.Parameters.AddWithValue("@CustomerFeedbackTypeID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete Customer Feedback Type
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
        [Route("ExportCustomerFeedbackTypeListToExcel")]
        public HttpResponseMessage ExportCustomerFeedbackTypeListToExcel()
        {
            try
            {
                string FileName = "CustomerFeedbackTypesList.xlsx";

                //Create a list to hold the list of Customer Feedback Types
                List<CustomerFeedbackType> customerFeedbackTypesList = new List<CustomerFeedbackType>();
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
                    cmd.CommandText = "spCustomerFeedbackType";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of Customer Feedback Types
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            CustomerFeedbackType customerFeedbackType = new CustomerFeedbackType();
                            customerFeedbackType.CustomerFeedbackTypeID = Convert.ToInt32(sqlReader["CustomerFeedbackTypeID"]);
                            customerFeedbackType.FeedbackType = sqlReader["FeedbackType"].ToString();
                            customerFeedbackType.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                            customerFeedbackTypesList.Add(customerFeedbackType);
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Customer Feedback Type ID");
                        ws.Cells[0, 2].PutValue("Feedback Type");
                        ws.Cells[0, 3].PutValue("Is Active?");

                        for (int c = 0; c <= 3; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (CustomerFeedbackType customerFeedbackType in customerFeedbackTypesList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(customerFeedbackType.CustomerFeedbackTypeID);
                            ws.Cells[row, 2].PutValue(customerFeedbackType.FeedbackType);
                            if (customerFeedbackType.IsActive)
                                ws.Cells[row, 3].PutValue("Yes");
                            else
                                ws.Cells[row, 3].PutValue("No");
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
