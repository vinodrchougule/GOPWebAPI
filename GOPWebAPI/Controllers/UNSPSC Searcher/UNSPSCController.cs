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
using System.Reflection;
using System.Web;
using System.Web.Http;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/UNSPSC")]
    public class UNSPSCController : ApiController
    {
        #region Read UNSPSC Versions
        [HttpGet]
        [Route("ReadUNSPSCVersions")]
        public IHttpActionResult ReadUNSPSCVersions()
        {
            try
            {
                //Create a list to hold the list of UNSPSC Versions
                List<UNSPSCVersions> UNSPSCVersionsList = new List<UNSPSCVersions>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUNSPSCGetVersions";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get list of UNSPSC Versions
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        UNSPSCVersions uNSPSCVersions = new UNSPSCVersions();
                        uNSPSCVersions.Version = sqlReader["Scheme_Name"].ToString();
                        UNSPSCVersionsList.Add(uNSPSCVersions);
                    }
                    conn.Close();

                    return Ok(UNSPSCVersionsList);
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

        #region Get UNSPSC Latest Version
        [HttpGet]
        [Route("GetUNSPSCLatestVersion")]
        public HttpResponseMessage GetUNSPSCLatestVersion()
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUNSPSCGetLatestVersion";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Calling sp to get the latest UNSPSC Version
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if(sqlReader.Read())
                        return Request.CreateResponse(HttpStatusCode.OK, sqlReader["Scheme_Name"].ToString());

                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Latest version not assigned");
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

        #region Read UNSPSC Search Result
        [HttpPost]
        [Route("ReadUNSPSCSearchResult")]
        public IHttpActionResult ReadUNSPSCSearchResult(UNSPSCSearchCriteriaModel model)
        {
            try
            {
                //Create a list to hold the list of search result
                List<UNSPSCSearchResultModel> searchResult = new List<UNSPSCSearchResultModel>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUNSPSCSearcherGetDataPagewise";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@SelectedVersion", model.TableName);
                    cmd.Parameters.AddWithValue("@Keyword1", model.Keyword1);
                    cmd.Parameters.AddWithValue("@Keyword2", model.Keyword2);
                    cmd.Parameters.AddWithValue("@IsAND", model.ANDStatus);
                    cmd.Parameters.AddWithValue("@IsOR", model.ORStatus);
                    cmd.Parameters.AddWithValue("@IsDoNotContain", model.DoNotContainStatus);
                    cmd.Parameters.AddWithValue("@UNSPSCCode1", model.UNSPSCCode1);
                    cmd.Parameters.AddWithValue("@UNSPSCCode2", model.UNSPSCCode2);
                    cmd.Parameters.AddWithValue("@UNSPSCCode3", model.UNSPSCCode3);
                    cmd.Parameters.AddWithValue("@UNSPSCCode4", model.UNSPSCCode4);
                    cmd.Parameters.AddWithValue("@PageNo", model.PageNo);
                    cmd.Parameters.AddWithValue("@PageSize", model.PageSize);
                    #endregion

                    //Calling sp to get the list of UNSPSC Search Result
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        UNSPSCSearchResultModel resultModel = new UNSPSCSearchResultModel();
                        resultModel.RowNum = Convert.ToInt32(sqlReader["ROWNUM"]);
                        resultModel.TotalCount = Convert.ToInt32(sqlReader["TotalCount"]);
                        resultModel.Code = sqlReader["Code"].ToString();
                        resultModel.Category = sqlReader["Category"].ToString();
                        searchResult.Add(resultModel);
                    }
                    conn.Close();

                    return Ok(searchResult);
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

        #region Read Segment, Family, Class, Commodity
        [HttpGet]
        [Route("ReadSegmentFamilyClassCommodity/{TableName}/{Code}")]
        public IHttpActionResult ReadSegmentFamilyClassCommodity(string TableName,string Code)
        {
            try
            {
                UNSPSCCategoryModel unspscCategoryModel = new UNSPSCCategoryModel();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUNSPSCSearcherGetCategories";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@SelectedVersion", TableName);
                    cmd.Parameters.AddWithValue("@Code", Code);
                    #endregion

                    //Calling sp to get the UNSPSC Categories
                    conn.Open();
                    System.Data.Common.DbDataReader sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if(sqlReader.Read())
                    {
                        unspscCategoryModel.SegmentCode = sqlReader["SegmentCode"].ToString();
                        unspscCategoryModel.Segment = sqlReader["Segment"].ToString();
                        unspscCategoryModel.FamilyCode = sqlReader["FamilyCode"].ToString();
                        unspscCategoryModel.Family = sqlReader["Family"].ToString();
                        unspscCategoryModel.ClassCode = sqlReader["ClassCode"].ToString();
                        unspscCategoryModel.Class = sqlReader["Class"].ToString();
                        unspscCategoryModel.CommodityCode = sqlReader["CommodityCode"].ToString();
                        unspscCategoryModel.Commodity = sqlReader["Commodity"].ToString();
                        unspscCategoryModel.CategoryDefinition = sqlReader["CategoryDefinition"].ToString();
                    }

                    return Ok(unspscCategoryModel);
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

        #region Download Help Document file
        [HttpGet]
        [Route("DownloadHelpDocument")]
        public HttpResponseMessage DownloadHelpDocument()
        {
            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string FileName = "UserGuideUNSPSCSearcher.pptx";

                //Set the Help File Path
                string filePath = HttpContext.Current.Server.MapPath("~/HelpDocs/") + FileName;

                //Check whether File exists.
                if (!File.Exists(filePath))
                    return Request.CreateErrorResponse(HttpStatusCode.NotFound, "File not found");

                //Read the File into a Byte Array.
                byte[] bytes = File.ReadAllBytes(filePath);

                //Set the Response Content.
                response.Content = new ByteArrayContent(bytes);

                //Set the Response Content Length.
                response.Content.Headers.ContentLength = bytes.LongLength;

                //Set the Content Disposition Header Value and FileName.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");

                response.Content.Headers.ContentDisposition.FileName = FileName;

                //Set the File Content Type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(FileName));

                return response;
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                HttpResponseMessage response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                return response;
            }
        }
        #endregion

        #region Read all categories of selected UNSPSC Version
        [HttpGet]
        [Route("ReadAllCategoriesOfSelecetdUNSPSCVersion")]
        public IHttpActionResult ReadAllCategoriesOfSelecetdUNSPSCVersion(string UNSPSCVersion)
        {
            try
            {
                //Create a list to hold the list of UNSPSC categories
                List<UNSPSCSearchResultModel> UNSPSCCategoriesList = new List<UNSPSCSearchResultModel>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spUNSPSCGetVersionAllCategories";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@UNSPSCVersion", UNSPSCVersion);
                    #endregion

                    //Calling sp to get the list of UNSPSC Categories result
                    conn.Open();
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        UNSPSCSearchResultModel resultModel = new UNSPSCSearchResultModel();
                        resultModel.Code = sqlReader["Code"].ToString();
                        resultModel.Category = sqlReader["Category"].ToString();
                        UNSPSCCategoriesList.Add(resultModel);
                    }
                    conn.Close();

                    return Ok(UNSPSCCategoriesList);
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
