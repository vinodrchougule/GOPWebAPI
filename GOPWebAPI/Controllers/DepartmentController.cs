using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Configuration;
using GOPWebAPI.Helpers;
using System.Web;
using Aspose.Cells;
using System.Net.Http.Headers;
using System.IO;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/department")]
    public class DepartmentController : ApiController
    {
        #region Read Departments
        [HttpGet]
        [Route]
        public IHttpActionResult ReadDepartments(bool IsToFetchOnlyOperationsDepartments=false)
        {
            try
            {
                int SlNo = 1;

                //Create a list to hold the list of departments
                List<Department> DepartmentsList = new List<Department>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spDepartment";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values - Mode - 4 - Read All
                    cmd.Parameters.AddWithValue("@Mode", DataMode.ReadAll);
                    cmd.Parameters.AddWithValue("@IsToFetchOnlyOperationsDepartments", IsToFetchOnlyOperationsDepartments);

                    //Calling sp to get list of departments
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        Department department = new Department();
                        department.SlNo = SlNo;
                        department.DepartmentID = Convert.ToInt32(sqlReader["DepartmentID"]);
                        department.Name = sqlReader["Name"].ToString();
                        department.IsActive = Convert.ToBoolean(sqlReader["IsActive"]);
                        DepartmentsList.Add(department);
                        SlNo++;
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(DepartmentsList);
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

        #region Read Department Head Count
        [HttpGet]
        [Route("ReadDepartmentHeadcount/{Department}")]
        public IHttpActionResult ReadDepartmentHeadcount(string Department="")
        {
            try
            {
                //Create a list to hold the list of department head count
                List<DepartmentHeadCount> DepartmentHeadCountList = new List<DepartmentHeadCount>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spDepartmentHeadCount";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@Department", Department);

                    //Calling sp to get list of department head count
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        DepartmentHeadCount departmentHeadCount = new DepartmentHeadCount();

                        departmentHeadCount.Name = sqlReader["Name"].ToString();
                        departmentHeadCount.HeadCount = Convert.ToInt32(sqlReader["HeadCount"]);
                        DepartmentHeadCountList.Add(departmentHeadCount);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(DepartmentHeadCountList);
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
