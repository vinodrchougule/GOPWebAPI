using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data;
using System.Data.SqlClient;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/postprojectbatchdetails")]
    public class PostProjectBatchDetailsController : ApiController
    {
        #region Create / Update Post Project Batch Details
        [HttpPost]
        [Route]
        public HttpResponseMessage UpdatePostProjectBatchDetails([FromBody]PostProjectBatchDetailsModel postProjectBatchDetailsModel)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Post Project Batch Details Data");

                if (!AccessControl.CanUserAccessPage(postProjectBatchDetailsModel.UserID, "Post Project Batch Details"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spPostProjectBatchDetails";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@CustomerCode", postProjectBatchDetailsModel.CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", postProjectBatchDetailsModel.ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", postProjectBatchDetailsModel.BatchNo);
                    cmd.Parameters.AddWithValue("@DuplicateCount", postProjectBatchDetailsModel.DuplicateCount);
                    cmd.Parameters.AddWithValue("@ExceptionalCount", postProjectBatchDetailsModel.ExceptionalCount);
                    cmd.Parameters.AddWithValue("@NotProcessedCount", postProjectBatchDetailsModel.NotProcessedCount);
                    cmd.Parameters.AddWithValue("@QCSamplingPercentage", postProjectBatchDetailsModel.QCSamplingPercentage);
                    cmd.Parameters.AddWithValue("@QCErrorRate", postProjectBatchDetailsModel.QCErrorRate);
                    cmd.Parameters.AddWithValue("@QASamplingPercentage", postProjectBatchDetailsModel.QASamplingPercentage);
                    cmd.Parameters.AddWithValue("@QAErrorRate", postProjectBatchDetailsModel.QAErrorRate);
                    cmd.Parameters.AddWithValue("@Remarks", postProjectBatchDetailsModel.Remarks);
                    cmd.Parameters.AddWithValue("@CAPADetails", postProjectBatchDetailsModel.CAPADetails);
                    cmd.Parameters.AddWithValue("@UserID", postProjectBatchDetailsModel.UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.Update);
                    #endregion

                    //Calling sp to create/update post project batch details
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

        #region Read Post Project Batch Details By Batch No.
        [HttpGet]
        [Route("{CustomerCode}/{ProjectCode}/{BatchNo}/{UserID}")]
        public IHttpActionResult ReadPostProjectBatchDetailsByBatchNo(string CustomerCode, string ProjectCode, string BatchNo, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Post Project Batch Details"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                PostProjectBatchDetailsModel postProjectBatchDetails = new PostProjectBatchDetailsModel();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spPostProjectBatchDetails";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 5 - Read By Id
                    cmd.Parameters.AddWithValue("@CustomerCode", CustomerCode);
                    cmd.Parameters.AddWithValue("@ProjectCode", ProjectCode);
                    cmd.Parameters.AddWithValue("@BatchNo", BatchNo);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadOne);

                    //Calling sp to get post project batch details
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        postProjectBatchDetails.CustomerCode = sqlReader["CustomerCode"].ToString();
                        postProjectBatchDetails.ProjectCode = sqlReader["ProjectCode"].ToString();
                        postProjectBatchDetails.BatchNo = sqlReader["BatchNo"].ToString();
                        postProjectBatchDetails.DuplicateCount = Convert.ToInt64(sqlReader["DuplicateCount"]);
                        postProjectBatchDetails.ExceptionalCount = Convert.ToInt64(sqlReader["ExceptionalCount"]);
                        postProjectBatchDetails.NotProcessedCount = Convert.ToInt64(sqlReader["NotProcessedCount"]);
                        postProjectBatchDetails.QCSamplingPercentage = Convert.ToDecimal(sqlReader["QCSamplingPercentage"]);
                        postProjectBatchDetails.QCErrorRate = Convert.ToDecimal(sqlReader["QCErrorRate"]);
                        postProjectBatchDetails.QASamplingPercentage = Convert.ToDecimal(sqlReader["QASamplingPercentage"]);
                        postProjectBatchDetails.QAErrorRate = Convert.ToDecimal(sqlReader["QAErrorRate"]);
                        postProjectBatchDetails.Remarks = sqlReader["Remarks"].ToString();
                        postProjectBatchDetails.CAPADetails = sqlReader["CAPADetails"].ToString();
                    }
                    conn.Close();

                    //return post project batch details to the request
                    return Ok(postProjectBatchDetails);
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
    }
}
