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
using System.Data.OleDb;
using System.Data.Common;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/ProjectsSummaryReport")]
    public class ProjectsSummaryReportController : ApiController
    {
        #region Read Projects Summary Report Data
        [HttpGet]
        [Route("ReadProjectsSummaryReportData/{FromDate}/{ToDate}")]
        public IHttpActionResult ReadProjectsSummaryReportData(DateTime FromDate, DateTime ToDate)
        {
            try
            {
                int index = 1;

                //Create a list to hold the Project Summary Report Header data
                List<ProjectsSummaryReportModel> ProjectsSummaryList = new List<ProjectsSummaryReportModel>();

                System.Data.Common.DbDataReader sqlReader, sqlProjectActivitiesReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spProjectsSummaryWithinDateRange";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@FromDate", FromDate);
                    cmd.Parameters.AddWithValue("@ToDate", ToDate);

                    //Initialize command object for fetching activity details
                    SqlCommand cmdActivities = conn.CreateCommand();
                    cmdActivities.CommandText = "spProjectStatusActivitywiseSummary";
                    cmdActivities.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of Project Report Data
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        ProjectsSummaryReportModel projectsSummaryReportModel = new ProjectsSummaryReportModel();

                        projectsSummaryReportModel.index = index;
                        projectsSummaryReportModel.ProjectID = Convert.ToInt64(sqlReader["ProjectID"]);
                        projectsSummaryReportModel.CustomerCode = sqlReader["CustomerCode"].ToString();
                        projectsSummaryReportModel.ProjectCode = sqlReader["ProjectCode"].ToString();
                        projectsSummaryReportModel.BatchNo = sqlReader["BatchNo"].ToString();
                        projectsSummaryReportModel.Scope = sqlReader["Scope"].ToString();
                        projectsSummaryReportModel.InputCount = Convert.ToInt64(sqlReader["InputCount"]);

                        ProjectsSummaryList.Add(projectsSummaryReportModel);
                        index++;
                    }
                    conn.Close();
                                                           
                    foreach (ProjectsSummaryReportModel projectsSummaryReportModel in ProjectsSummaryList)
                    {

                        DataTable dtActivityDetails = new DataTable();
                        dtActivityDetails.Columns.Add("Activity");
                        dtActivityDetails.Columns.Add("ActivityCount");
                        dtActivityDetails.Columns.Add("ProductionCompletedCount");
                        dtActivityDetails.Columns.Add("ProductionCompletedPercentage");
                        dtActivityDetails.Columns.Add("QCCompletedCount");
                        dtActivityDetails.Columns.Add("QCCompletedPercentage");

                        //Add parameters with values
                        cmdActivities.Parameters.Clear();
                        cmdActivities.Parameters.AddWithValue("@CustomerCode", projectsSummaryReportModel.CustomerCode);
                        cmdActivities.Parameters.AddWithValue("@ProjectCode", projectsSummaryReportModel.ProjectCode);
                        cmdActivities.Parameters.AddWithValue("@BatchNo", projectsSummaryReportModel.BatchNo);

                        conn.Open();
                        cmdActivities.CommandTimeout = 0;
                        sqlProjectActivitiesReader = (System.Data.Common.DbDataReader)cmdActivities.ExecuteReader();
                        while (sqlProjectActivitiesReader.Read())
                        {
                            DataRow dr = dtActivityDetails.NewRow();

                            dr["Activity"] = sqlProjectActivitiesReader["Activity"].ToString();
                            dr["ActivityCount"]=Convert.ToInt32(sqlProjectActivitiesReader["ActivityCount"]);
                            dr["ProductionCompletedCount"]=Convert.ToInt32(sqlProjectActivitiesReader["ProductionCompletedCount"]);
                            dr["ProductionCompletedPercentage"]=Convert.ToDecimal(sqlProjectActivitiesReader["ProductionCompletedPercentage"]);
                            dr["QCCompletedCount"]=Convert.ToInt32(sqlProjectActivitiesReader["QCCompletedCount"]);
                            dr["QCCompletedPercentage"]=Convert.ToDecimal(sqlProjectActivitiesReader["QCCompletedPercentage"]);

                            dtActivityDetails.Rows.Add(dr);
                        }

                        ProjectsSummaryList.Where(psl => psl.CustomerCode == projectsSummaryReportModel.CustomerCode &&
                                                         psl.ProjectCode == projectsSummaryReportModel.ProjectCode &&
                                                         psl.BatchNo == projectsSummaryReportModel.BatchNo).First().ActivityDetails = dtActivityDetails;

                        conn.Close();
                    }

                    //return list to the request
                    return Ok(ProjectsSummaryList);
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
