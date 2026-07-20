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
using System.Text.RegularExpressions;
using System.Web.Routing;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/snomedsearch")]
    public class SNOMEDSearchController : ApiController
    {
        #region Read Concept Terms
        [HttpPost]
        [Route("ReadConceptTerms")]
        public IHttpActionResult ReadConceptTerms([FromBody] SNOMEDSearchModel sNOMEDSearchModel)
        {
            try
            {
                List<SNOMEDConceptTerm> SNOMEDConceptTermsList = new List<SNOMEDConceptTerm>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSNOMEDGetConceptsAndTerms";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@SearchText1", sNOMEDSearchModel.SearchText1);
                    cmd.Parameters.AddWithValue("@SearchText2", sNOMEDSearchModel.SearchText2);
                    cmd.Parameters.AddWithValue("@SearchText3", sNOMEDSearchModel.SearchText3);
                    cmd.Parameters.AddWithValue("@SearchText4", sNOMEDSearchModel.SearchText4);

                    //Calling sp to get list of Concept Terms
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        SNOMEDConceptTerm sNOMEDConceptTerm = new SNOMEDConceptTerm();
                        sNOMEDConceptTerm.Active = sqlReader["active"].ToString();
                        sNOMEDConceptTerm.IsFSN = sqlReader["Is_FSN"].ToString();
                        sNOMEDConceptTerm.ConceptId = sqlReader["ConceptId"].ToString();
                        sNOMEDConceptTerm.Term = sqlReader["Term"].ToString();

                        SNOMEDConceptTermsList.Add(sNOMEDConceptTerm);
                    }
                    conn.Close();
                }

                //return list to the request
                return Ok(SNOMEDConceptTermsList);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Concept Synonyms
        [HttpGet]
        [Route("ReadConceptSynonyms/{ConceptID}")]
        public IHttpActionResult ReadConceptSynonyms(string ConceptID)
        {
            try
            {
                List<SNOMEDConceptSynonym> SNOMEDConceptSynonymList = new List<SNOMEDConceptSynonym>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSNOMEDGetConceptSynonyms";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ConceptId", ConceptID);

                    //Calling sp to get list of Concept Synonyms
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        SNOMEDConceptSynonym sNOMEDConceptSynonym = new SNOMEDConceptSynonym();
                        sNOMEDConceptSynonym.ConceptId = sqlReader["ConceptId"].ToString();
                        sNOMEDConceptSynonym.Term = sqlReader["Term"].ToString();

                        SNOMEDConceptSynonymList.Add(sNOMEDConceptSynonym);
                    }
                    conn.Close();
                }

                //return list to the request
                return Ok(SNOMEDConceptSynonymList);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Concept Parents
        [HttpGet]
        [Route("ReadConceptParents/{ConceptID}")]
        public IHttpActionResult ReadConceptParents(string ConceptID)
        {
            try
            {
                List<SNOMEDConceptParent> SNOMEDConceptParentsList = new List<SNOMEDConceptParent>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSNOMEDGetConceptParents";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ConceptId", ConceptID);

                    //Calling sp to get list of Concept Parents
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        SNOMEDConceptParent sNOMEDConceptParent = new SNOMEDConceptParent();
                        sNOMEDConceptParent.DestinationID = sqlReader["DestinationID"].ToString();
                        sNOMEDConceptParent.DestinationIDName = sqlReader["DestinationIDName"].ToString();

                        SNOMEDConceptParentsList.Add(sNOMEDConceptParent);
                    }
                    conn.Close();
                }

                //return list to the request
                return Ok(SNOMEDConceptParentsList);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Content(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read Concept Children
        [HttpGet]
        [Route("ReadConceptChildren/{ConceptID}")]
        public IHttpActionResult ReadConceptChildren(string ConceptID)
        {
            try
            {
                List<SNOMEDConceptChildren> SNOMEDConceptChildrenList = new List<SNOMEDConceptChildren>();

                System.Data.Common.DbDataReader sqlReader;
                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spSNOMEDGetConceptChildren";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    //Add parameters with values
                    cmd.Parameters.AddWithValue("@ConceptId", ConceptID);

                    //Calling sp to get list of Concept Children
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        SNOMEDConceptChildren sNOMEDConceptChildren = new SNOMEDConceptChildren();
                        sNOMEDConceptChildren.SourceId = sqlReader["SourceId"].ToString();
                        sNOMEDConceptChildren.SourceIDName = sqlReader["SourceIDName"].ToString();

                        SNOMEDConceptChildrenList.Add(sNOMEDConceptChildren);
                    }
                    conn.Close();
                }

                //return list to the request
                return Ok(SNOMEDConceptChildrenList);
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

                string FileName = "UserGuideSNOMEDSearcher.pptx";

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
    }
}
