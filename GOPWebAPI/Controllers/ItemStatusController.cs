using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using context = System.Web.HttpContext;
using System.Configuration;
using GOPWebAPI.Helpers;
using System.Web;
using Aspose.Cells;
using System.IO;
using System.Net.Http.Headers;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/itemstatus")]
    public class ItemStatusController : ApiController
    {
        #region Create Item Status
        [HttpPost]
        [Route]
        public HttpResponseMessage CreateItemStatus([FromBody]ItemStatus itemStatus)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                if (!AccessControl.CanUserAccessPage(itemStatus.UserID, "Create Item Status"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spItemStatus";

                    //Add parameters with values - Mode - 1 - Create
                    cmd.Parameters.AddWithValue("@Status", itemStatus.Status);
                    cmd.Parameters.AddWithValue("@IsActive", itemStatus.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", itemStatus.UserID);
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

        #region Read Item Status List
        [HttpGet]
        [Route("{UserID}")]
        public IHttpActionResult ReadItemStatusList(string UserID)
        {
            try
            {
                int SlNo = 1;

                if (!AccessControl.CanUserAccessPage(UserID, "Item Status List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the list of Item Status
                List<ItemStatus> ItemStatusList = new List<ItemStatus>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spItemStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of Item Status
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ItemStatus itemStatus = new ItemStatus();
                        itemStatus.SlNo = SlNo;
                        itemStatus.ItemStatusID = Convert.ToInt32(sqlReader["ItemStatusID"]);
                        itemStatus.Status = sqlReader["Status"].ToString();
                        itemStatus.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        ItemStatusList.Add(itemStatus);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(ItemStatusList);
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

        #region Read Item Status By Id
        [HttpGet]
        [Route("{id}/{UserID}")]
        public IHttpActionResult ReadItemStatusById(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Item Status"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create item status instance
                ItemStatus itemStatus = new ItemStatus();
                System.Data.Common.DbDataReader sqlReader;
                
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spItemStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@ItemStatusID", id);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get item status details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        itemStatus.ItemStatusID = Convert.ToInt32(sqlReader["ItemStatusID"]);
                        itemStatus.Status = sqlReader["Status"].ToString();
                        itemStatus.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        conn.Close();

                        //return item status to the request
                        return Ok(itemStatus);
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

        #region Update Item Status
        [HttpPut]
        [Route("{id}")]
        public HttpResponseMessage UpdateItemStatus(int id, [FromBody]ItemStatus itemStatus)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                //Check Item Status Id
                if (id != itemStatus.ItemStatusID)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Item Status id");

                if (!AccessControl.CanUserAccessPage(itemStatus.UserID, "Edit Item Status"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spItemStatus";

                    //Add parameters with values - Mode - 2 - Update
                    cmd.Parameters.AddWithValue("@ItemStatusID", itemStatus.ItemStatusID);
                    cmd.Parameters.AddWithValue("@Status", itemStatus.Status);
                    cmd.Parameters.AddWithValue("@IsActive", itemStatus.IsActive);
                    cmd.Parameters.AddWithValue("@UserID", itemStatus.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);

                    //Calling sp to update Item Status
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

        #region Delete Item Status
        [HttpPatch]
        [Route("{id}/{UserID}")]
        public HttpResponseMessage DeleteItemStatus(int id, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Delete Item Status"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spItemStatus";

                    //Add parameters with values - Mode - 3 - Delete
                    cmd.Parameters.AddWithValue("@ItemStatusID", id);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Delete);

                    //Calling sp to delete Item Status
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
        [Route("ExportItemStatusListToExcel")]
        public HttpResponseMessage ExportItemStatusListToExcel()
        {
            try
            {
                string FileName = "ItemStatusList.xlsx";

                //Create a list to hold the list of Item Status
                List<ItemStatus> itemStatusList = new List<ItemStatus>();
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
                    cmd.CommandText = "spItemStatus";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);

                    //Calling sp to get list of Item Status
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.HasRows)
                    {
                        while (sqlReader.Read())
                        {
                            ItemStatus itemStatus = new ItemStatus();
                            itemStatus.ItemStatusID = Convert.ToInt32(sqlReader["ItemStatusID"]);
                            itemStatus.Status = sqlReader["Status"].ToString();
                            itemStatus.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                            itemStatusList.Add(itemStatus);
                        }
                        conn.Close();

                        #region Writing column headings and setting style
                        ws.Cells[0, 0].PutValue("S.No.");
                        ws.Cells[0, 1].PutValue("Item Status ID");
                        ws.Cells[0, 2].PutValue("Status");
                        ws.Cells[0, 3].PutValue("Is Active?");

                        for (int c = 0; c <= 3; c++)
                            ws.Cells[0, c].SetStyle(styleHeader);
                        #endregion

                        #region Writing row data
                        foreach (ItemStatus itemStatus in itemStatusList)
                        {
                            #region Writing row data
                            ws.Cells[row, 0].PutValue(row);
                            ws.Cells[row, 1].PutValue(itemStatus.ItemStatusID);
                            ws.Cells[row, 2].PutValue(itemStatus.Status);
                            if (itemStatus.IsActive)
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
