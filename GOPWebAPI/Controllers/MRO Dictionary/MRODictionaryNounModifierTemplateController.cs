using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Xml.Serialization;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models.MRO_Dictionary;
using System.Xml;
using System.Reflection;
using Microsoft.Ajax.Utilities;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;

namespace GOPWebAPI.Controllers.MRO_Dictionary
{
    [RoutePrefix("api/MRODictionaryNounModifierTemplate")]
    public class MRODictionaryNounModifierTemplateController : ApiController
    {
        #region Create Noun-Modifier Template
        [HttpPost]
        [Route("CreateNounModifierTemplate")]
        public HttpResponseMessage CreateNounModifierTemplate([FromBody] MRODictionaryNounModifierTemplateModel model)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Noun-Modifier Template Data");

                if (!AccessControl.CanUserAccessPage(model.UserID, "Create Noun-Modifier Template"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Noun Synonyms
                //create xml serialized data
                var serializerNounSynonym = new XmlSerializer(typeof(List<NounSynonym>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounSynonym = new StringWriter();

                //write serialized data to stream
                serializerNounSynonym.Serialize(streamNounSynonym, model.NounSynonyms);

                //Read stream as xml string
                StringReader transactionXmlNounSynonym = new StringReader(streamNounSynonym.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounSynonym = new XmlTextReader(transactionXmlNounSynonym);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounSynonym = new SqlXml(xmlReaderNounSynonym);
                #endregion

                #region Noun-Modifier Attributes
                //create xml serialized data
                var serializerNounModifierTemplateAttribute = new XmlSerializer(typeof(List<NounModifierTemplateAttribute>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounModifierTemplateAttribute = new StringWriter();

                //write serialized data to stream
                serializerNounModifierTemplateAttribute.Serialize(streamNounModifierTemplateAttribute, model.NounModifierAttributes);

                //Read stream as xml string
                StringReader transactionXmlNounModifierTemplateAttribute = new StringReader(streamNounModifierTemplateAttribute.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounModifierTemplateAttribute = new XmlTextReader(transactionXmlNounModifierTemplateAttribute);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounModifierTemplateAttribute = new SqlXml(xmlReaderNounModifierTemplateAttribute);
                #endregion

                #region Noun-Modifier Attribute Values
                //create xml serialized data
                var serializerNounModifierTemplateAttributeValues = new XmlSerializer(typeof(List<NounModifierTemplateAttributeEVV>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounModifierTemplateAttributeValues = new StringWriter();

                //write serialized data to stream
                serializerNounModifierTemplateAttributeValues.Serialize(streamNounModifierTemplateAttributeValues, model.NounModifierAttributeEVVs);

                //Read stream as xml string
                StringReader transactionXmlNounModifierTemplateAttributeValues = new StringReader(streamNounModifierTemplateAttributeValues.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounModifierTemplateAttributeValues = new XmlTextReader(transactionXmlNounModifierTemplateAttributeValues);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounModifierTemplateAttributeValues = new SqlXml(xmlReaderNounModifierTemplateAttributeValues);
                #endregion

                #region Noun-Modifier UNSPSCs
                //create xml serialized data
                var serializerNounModifierUNSPSC = new XmlSerializer(typeof(List<NounModifierUNSPSC>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounModifierUNSPSC = new StringWriter();

                //write serialized data to stream
                serializerNounModifierUNSPSC.Serialize(streamNounModifierUNSPSC, model.NounModifierUNSPSCs);

                //Read stream as xml string
                StringReader transactionXmlNounModifierUNSPSC = new StringReader(streamNounModifierUNSPSC.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounModifierUNSPSC = new XmlTextReader(transactionXmlNounModifierUNSPSC);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounModifierUNSPSC = new SqlXml(xmlReaderNounModifierUNSPSC);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spMRODictionaryCreateNMTemplate";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", model.VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", model.Noun);
                    cmd.Parameters.AddWithValue("@Modifier", model.Modifier);
                    cmd.Parameters.AddWithValue("@NounDefinition", model.NounDefinition);
                    cmd.Parameters.AddWithValue("@NounModifierDefinitionOrGuidelines", model.NounModifierDefinitionOrGuidelines);
                    cmd.Parameters.Add(new SqlParameter("@NounModifierSynonyms", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounSynonym
                    });

                    cmd.Parameters.Add(new SqlParameter("@NounModifierAttributes", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounModifierTemplateAttribute
                    });

                    cmd.Parameters.Add(new SqlParameter("@NounModifierAttributeEVVs", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounModifierTemplateAttributeValues
                    });

                    cmd.Parameters.Add(new SqlParameter("@NounModifierUNSPSC", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounModifierUNSPSC
                    });
                    cmd.Parameters.AddWithValue("@UserID", model.UserID);
                    #endregion

                    //Calling sp to create Noun-Modifier Template
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "created")
                    {
                        #region Moving Noun-Modifier Images from temp folder to Uploads folder
                        DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/MRODictionary/ProductImages/"));

                        int imageCounter = 1;
                        string imageFileExtension = string.Empty,imageNewFileName = string.Empty;

                        foreach(string fileName in model.ImageFileNames)
                        {
                            if (File.Exists(dirTemp + fileName))
                            {
                                imageFileExtension = Path.GetExtension(dirTemp + fileName);
                                imageNewFileName = model.Noun.Trim().ToUpper() + "_" + model.Modifier.Trim().ToUpper() + imageCounter.ToString() + imageFileExtension;
                                FileOperations.MoveFile(dirTemp, dirTemp + fileName, dirUploads, imageNewFileName);
                                imageCounter++;
                            }
                        }
                        #endregion

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

        #region Read Noun-Modifier Images from Uploads folder
        [HttpPost]
        [Route("ReadNounModifierImages")]
        //IsOnlyToView flag to be passed as false while editing Noun-Modifier template
        public HttpResponseMessage ReadNounModifierImages(string Noun, string Modifier, bool IsOnlyToView = true)
        {
            List<ImageModel> images = new List<ImageModel>();

            //Set the Image Folder Path.
            string imagesPath = HttpContext.Current.Server.MapPath("~/Uploads/MRODictionary/ProductImages");
            string imagesTempPath = HttpContext.Current.Server.MapPath("~/temp/");
            string imageTempFileName;

            foreach (string file in Directory.GetFiles(imagesPath))
            {
                FileInfo fileInfo = new FileInfo(file);

                string fileNameWithoutExtension =fileInfo.Name.Substring(0, fileInfo.Name.LastIndexOf('.'));
                string fileNameWithoutTrailingDigits = Regex.Replace(fileNameWithoutExtension, @"[\d-]", string.Empty);
                string[] fileNameSplittedArray = fileNameWithoutTrailingDigits.Split('_');

                //if (fileInfo.Name.Trim().ToUpper().StartsWith(Noun.Trim().ToUpper() + "_" + Modifier.Trim().ToUpper()))
                if (fileNameSplittedArray.Length == 2)
                {
                    if (fileNameSplittedArray[0].Trim().ToUpper() == Noun.Trim().ToUpper() &&
                        fileNameSplittedArray[1].Trim().ToUpper() == Modifier.Trim().ToUpper())
                    {
                        //Read the Image as Byte Array.
                        byte[] bytes = File.ReadAllBytes(fileInfo.FullName);

                        imageTempFileName = Guid.NewGuid() + Path.GetExtension(fileInfo.Name);
                        images.Add(new ImageModel
                        {
                            Name = Path.GetFileName(fileInfo.Name),
                            Data = Convert.ToBase64String(bytes, 0, bytes.Length),
                            ImageTempFileName = imageTempFileName
                        });

                        if (!IsOnlyToView)
                            File.Copy(fileInfo.FullName, imagesTempPath + imageTempFileName, true);
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, images);
        }
        #endregion

        #region Read all images uploaded from Temp folder
        [HttpPost]
        [Route("ReadAllImagesFromTempFolder")]
        public HttpResponseMessage ReadAllImagesFromTempFolder(List<string> imageFileNamesList)
        {
            List<ImageModel> images = new List<ImageModel>();

            //Set the Image Folder Path.
            string imagesTempPath = HttpContext.Current.Server.MapPath("~/temp/");

            if (imageFileNamesList != null)
            {
                foreach (string tempFileName in imageFileNamesList)
                {
                    foreach (string file in Directory.GetFiles(imagesTempPath))
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Name.Trim().ToUpper() == tempFileName.Trim().ToUpper())
                        {
                            //Read the Image as Byte Array.
                            byte[] bytes = File.ReadAllBytes(fileInfo.FullName);

                            images.Add(new ImageModel
                            {
                                Name = Path.GetFileName(fileInfo.Name),
                                Data = Convert.ToBase64String(bytes, 0, bytes.Length),
                                ImageTempFileName = tempFileName
                            });
                        }
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, images);
        }
        #endregion

        #region Read Noun-Modifiers Template list
        [HttpGet]
        [Route("ReadNounModifiersTemplateList")]
        public IHttpActionResult ReadNounModifiersTemplateList(string UserID, string VersionNameOrNo="")
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "Noun-Modifier Template List"))
                    return Content(HttpStatusCode.BadRequest, "Access Denied");

                //Create a list to hold the MRO Dictionary Noun-Modifier List
                List<MRODictionaryNounModifier> mroDictionaryNounModifiersList = new List<MRODictionaryNounModifier>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryNounModifiersList";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
                    #endregion

                    //Calling sp to get list of MRO Dictionary Noun-Modifiers Template List
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        MRODictionaryNounModifier mroDictionaryNounModifier = new MRODictionaryNounModifier();
                        mroDictionaryNounModifier.VersionNameOrNo = sqlReader["VersionNameOrNo"].ToString();
                        mroDictionaryNounModifier.Noun = sqlReader["Noun"].ToString();
                        mroDictionaryNounModifier.Modifier = sqlReader["Modifier"].ToString();
                        mroDictionaryNounModifier.CreatedOn = Convert.ToDateTime(sqlReader["CreatedOn"]);
                        mroDictionaryNounModifier.CreatedBy = sqlReader["CreatedBy"].ToString();
                        mroDictionaryNounModifier.UpdatedOn = sqlReader["UpdatedOn"] == DBNull.Value ? null : (DateTime?)sqlReader["UpdatedOn"];
                        mroDictionaryNounModifier.UpdatedBy = sqlReader["UpdatedBy"].ToString();
                        mroDictionaryNounModifiersList.Add(mroDictionaryNounModifier);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(mroDictionaryNounModifiersList);
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

        #region Read Noun-Modifier Details from selected Version
        [HttpGet]
        [Route("ReadNounModifierDetailsFromSelectedVersion")]
        public IHttpActionResult ReadNounModifierDetailsFromSelectedVersion(string VersionNameOrNo, string Noun, string Modifier)
        {
            try
            {
                System.Data.Common.DbDataReader sqlReader;
                MRODictionaryNounModifier mroDictionaryNounModifier = new MRODictionaryNounModifier();

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryVersionNounModifierDetailsRead";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);
                    #endregion

                    //Calling sp to get the MRO Dictionary Noun-Modifier details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    if (sqlReader.Read())
                    {
                        mroDictionaryNounModifier.VersionNameOrNo = VersionNameOrNo;
                        mroDictionaryNounModifier.Noun = Noun;
                        mroDictionaryNounModifier.Modifier = Modifier;
                        mroDictionaryNounModifier.NounDefinition = sqlReader["NounDefinition"].ToString();
                        mroDictionaryNounModifier.NounModifierDefinitionOrGuidelines = sqlReader["NounModifierDefinitionOrGuidelines"].ToString();
                        conn.Close();

                        //return object to the request
                        return Ok(mroDictionaryNounModifier);
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

        #region Read Noun Synonym Details from selected Version
        [HttpGet]
        [Route("ReadNounSynonymDetailsFromSelectedVersion")]
        public IHttpActionResult ReadNounSynonymDetailsFromSelectedVersion(string VersionNameOrNo, string Noun)
        {
            try
            {
                //Create a list to hold the MRO Dictionary Noun Synonym details
                List<NounSynonym> nounSynonymDetailsList = new List<NounSynonym>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryVersionNounSynonymDetailsRead";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    #endregion

                    //Calling sp to get list of Noun Synonym details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        NounSynonym nounSynonym = new NounSynonym();

                        nounSynonym.Synonym = sqlReader["Synonym"].ToString();
                        nounSynonym.SynonymDefinitionOrGuidelines = sqlReader["SynonymDefinitionOrGuidelines"].ToString();
                        nounSynonymDetailsList.Add(nounSynonym);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(nounSynonymDetailsList);
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

        #region Read Noun-Modifier Attribute Details from selected Version
        [HttpGet]
        [Route("ReadNounModifierAttributeDetailsFromSelectedVersion")]
        public IHttpActionResult ReadNounModifierAttributeDetailsFromSelectedVersion(string VersionNameOrNo, string Noun, string Modifier)
        {
            try
            {
                //Create a list to hold the MRO Dictionary Noun-Modifier Attribute details
                List<NounModifierTemplateAttribute> nounModifierAttributeDetailsList = new List<NounModifierTemplateAttribute>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryVersionNounModifierAttributeDetailsRead";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);
                    #endregion

                    //Calling sp to get list of Noun-Modifier Attribute details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        NounModifierTemplateAttribute nounModifierAttribute = new NounModifierTemplateAttribute();
                        nounModifierAttribute.Attribute = sqlReader["Attribute"].ToString();
                        nounModifierAttribute.Priority = sqlReader["Priority"].ToString();
                        nounModifierAttribute.MandatoryOrOptional = sqlReader["MandatoryOrOptional"].ToString();
                        nounModifierAttribute.AttributeGuidelines = sqlReader["AttributeGuidelines"].ToString();
                        nounModifierAttributeDetailsList.Add(nounModifierAttribute);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(nounModifierAttributeDetailsList);
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

        #region Read Noun-Modifier Attribute Values Details from selected Version
        [HttpGet]
        [Route("ReadNounModifierAttributeValuesDetailsFromSelectedVersion")]
        public IHttpActionResult ReadNounModifierAttributeValuesDetailsFromSelectedVersion(string VersionNameOrNo, string Noun, string Modifier)
        {
            try
            {
                //Create a list to hold the MRO Dictionary Noun-Modifier Attribute Values details
                List<NounModifierTemplateAttributeEVV> nounModifierAttributeValuesDetailsList = new List<NounModifierTemplateAttributeEVV>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryVersionNounModifierAttributeValuesDetailsRead";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);
                    #endregion

                    //Calling sp to get list of Noun-Modifier Attribute Values details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        NounModifierTemplateAttributeEVV nounModifierAttributeEVV = new NounModifierTemplateAttributeEVV();
                        nounModifierAttributeEVV.Attribute = sqlReader["Attribute"].ToString();
                        nounModifierAttributeEVV.EnumeratedValidValue = sqlReader["EnumeratedValidValue"].ToString();
                        nounModifierAttributeEVV.Priority = sqlReader["Priority"].ToString();
                        nounModifierAttributeValuesDetailsList.Add(nounModifierAttributeEVV);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(nounModifierAttributeValuesDetailsList);
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

        #region Read Noun-Modifier UNSPSC Details from selected Version
        [HttpGet]
        [Route("ReadNounModifierUNSPSCDetailsFromSelectedVersion")]
        public IHttpActionResult ReadNounModifierUNSPSCDetailsFromSelectedVersion(string VersionNameOrNo, string Noun, string Modifier)
        {
            try
            {
                //Create a list to hold the MRO Dictionary Noun-Modifier UNSPSC details
                List<NounModifierUNSPSC> nounModifierUNSPSCDetailsList = new List<NounModifierUNSPSC>();
                System.Data.Common.DbDataReader sqlReader;

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "spMRODictionaryVersionNounModifierUNSPSCDetailsRead";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);
                    #endregion

                    //Calling sp to get list of Noun-Modifier UNSPSC details
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    sqlReader = (System.Data.Common.DbDataReader)cmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        NounModifierUNSPSC nounModifierUNSPSC = new NounModifierUNSPSC();
                        nounModifierUNSPSC.UNSPSCVersion = sqlReader["Scheme_Name"].ToString();
                        nounModifierUNSPSC.UNSPSCCode = sqlReader["Code"].ToString();
                        nounModifierUNSPSC.UNSPSCCategory = sqlReader["Category"].ToString();
                        nounModifierUNSPSCDetailsList.Add(nounModifierUNSPSC);
                    }
                    conn.Close();

                    //return list to the request
                    return Ok(nounModifierUNSPSCDetailsList);
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

        #region Edit Noun-Modifier Template
        [HttpPost]
        [Route("EditNounModifierTemplate")]
        public HttpResponseMessage EditNounModifierTemplate([FromBody] MRODictionaryNounModifierTemplateModel model)
        {
            try
            {
                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Noun-Modifier Template Data");

                if (!AccessControl.CanUserAccessPage(model.UserID, "Edit Noun-Modifier Template"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                #region Noun Synonyms
                //create xml serialized data
                var serializerNounSynonym = new XmlSerializer(typeof(List<NounSynonym>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounSynonym = new StringWriter();

                //write serialized data to stream
                serializerNounSynonym.Serialize(streamNounSynonym, model.NounSynonyms);

                //Read stream as xml string
                StringReader transactionXmlNounSynonym = new StringReader(streamNounSynonym.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounSynonym = new XmlTextReader(transactionXmlNounSynonym);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounSynonym = new SqlXml(xmlReaderNounSynonym);
                #endregion

                #region Noun-Modifier Attributes
                //create xml serialized data
                var serializerNounModifierTemplateAttribute = new XmlSerializer(typeof(List<NounModifierTemplateAttribute>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounModifierTemplateAttribute = new StringWriter();

                //write serialized data to stream
                serializerNounModifierTemplateAttribute.Serialize(streamNounModifierTemplateAttribute, model.NounModifierAttributes);

                //Read stream as xml string
                StringReader transactionXmlNounModifierTemplateAttribute = new StringReader(streamNounModifierTemplateAttribute.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounModifierTemplateAttribute = new XmlTextReader(transactionXmlNounModifierTemplateAttribute);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounModifierTemplateAttribute = new SqlXml(xmlReaderNounModifierTemplateAttribute);
                #endregion

                #region Noun-Modifier Attribute Values
                //create xml serialized data
                var serializerNounModifierTemplateAttributeValues = new XmlSerializer(typeof(List<NounModifierTemplateAttributeEVV>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounModifierTemplateAttributeValues = new StringWriter();

                //write serialized data to stream
                serializerNounModifierTemplateAttributeValues.Serialize(streamNounModifierTemplateAttributeValues, model.NounModifierAttributeEVVs);

                //Read stream as xml string
                StringReader transactionXmlNounModifierTemplateAttributeValues = new StringReader(streamNounModifierTemplateAttributeValues.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounModifierTemplateAttributeValues = new XmlTextReader(transactionXmlNounModifierTemplateAttributeValues);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounModifierTemplateAttributeValues = new SqlXml(xmlReaderNounModifierTemplateAttributeValues);
                #endregion

                #region Noun-Modifier UNSPSCs
                //create xml serialized data
                var serializerNounModifierUNSPSC = new XmlSerializer(typeof(List<NounModifierUNSPSC>),
                                       new XmlRootAttribute("root"));
                //create a stream
                var streamNounModifierUNSPSC = new StringWriter();

                //write serialized data to stream
                serializerNounModifierUNSPSC.Serialize(streamNounModifierUNSPSC, model.NounModifierUNSPSCs);

                //Read stream as xml string
                StringReader transactionXmlNounModifierUNSPSC = new StringReader(streamNounModifierUNSPSC.ToString());

                //Read xml string
                XmlTextReader xmlReaderNounModifierUNSPSC = new XmlTextReader(transactionXmlNounModifierUNSPSC);

                //Convert xml string to sql xml
                SqlXml sqlXmlNounModifierUNSPSC = new SqlXml(xmlReaderNounModifierUNSPSC);
                #endregion

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spMRODictionaryEditNMTemplate";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", model.VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", model.Noun);
                    cmd.Parameters.AddWithValue("@Modifier", model.Modifier);
                    cmd.Parameters.AddWithValue("@NounDefinition", model.NounDefinition);
                    cmd.Parameters.AddWithValue("@NounModifierDefinitionOrGuidelines", model.NounModifierDefinitionOrGuidelines);
                    cmd.Parameters.Add(new SqlParameter("@NounModifierSynonyms", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounSynonym
                    });

                    cmd.Parameters.Add(new SqlParameter("@NounModifierAttributes", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounModifierTemplateAttribute
                    });

                    cmd.Parameters.Add(new SqlParameter("@NounModifierAttributeEVVs", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounModifierTemplateAttributeValues
                    });

                    cmd.Parameters.Add(new SqlParameter("@NounModifierUNSPSC", SqlDbType.Xml)
                    {
                        Value = sqlXmlNounModifierUNSPSC
                    });
                    cmd.Parameters.AddWithValue("@UserID", model.UserID);
                    #endregion

                    //Calling sp to edit Noun-Modifier Template
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "updated")
                    {
                        #region Deleting all Noun-Modifier Images from Uploads folder
                        string imagesUploadsPath = HttpContext.Current.Server.MapPath("~/Uploads/MRODictionary/ProductImages/");

                        foreach (string file in Directory.GetFiles(imagesUploadsPath))
                        {
                            FileInfo fileInfo = new FileInfo(file);

                            if (fileInfo.Name.Trim().ToUpper().StartsWith(model.Noun.Trim().ToUpper() + "_" + model.Modifier.Trim().ToUpper()))
                                File.Delete(fileInfo.FullName);
                        }
                        #endregion

                        #region Moving Noun-Modifier Images from temp folder to Uploads folder
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/MRODictionary/ProductImages/"));
                        DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                        int imageCounter = 1;
                        string imageFileExtension = string.Empty, imageNewFileName = string.Empty;

                        foreach (string fileName in model.ImageFileNames)
                        {
                            if (File.Exists(dirTemp + fileName))
                            {
                                imageFileExtension = Path.GetExtension(dirTemp + fileName);
                                imageNewFileName = model.Noun.Trim().ToUpper() + "_" + model.Modifier.Trim().ToUpper() + imageCounter.ToString() + imageFileExtension;
                                FileOperations.MoveFile(dirTemp, dirTemp + fileName, dirUploads, imageNewFileName);
                                imageCounter++;
                            }
                        }
                        #endregion

                        return Request.CreateResponse(HttpStatusCode.OK,Result);
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

        #region Delete Noun-Modifier Template
        [HttpPatch]
        [Route("DeleteNounModifierTemplate")]
        public HttpResponseMessage DeleteNounModifierTemplate(string VersionNameOrNo, string Noun, string Modifier, string UserID)
        {
            try
            {
                if(string.IsNullOrEmpty(VersionNameOrNo) || string.IsNullOrEmpty(Noun) || string.IsNullOrEmpty(Modifier))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid VersionNameOrNo and/or Noun and/or Modifier Template Data");

                if (!AccessControl.CanUserAccessPage(UserID, "Delete Noun-Modifier Template"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                using (SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString()))
                {
                    //Initialize command object
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "spMRODictionaryDeleteNMTemplate";

                    #region Adding Stored Procedure Parameters
                    cmd.Parameters.AddWithValue("@VersionNameOrNo", VersionNameOrNo);
                    cmd.Parameters.AddWithValue("@Noun", Noun);
                    cmd.Parameters.AddWithValue("@Modifier", Modifier);
                    cmd.Parameters.AddWithValue("@UserID", UserID);
                    #endregion

                    //Calling sp to delete Noun-Modifier Template
                    conn.Open();
                    cmd.CommandTimeout = 0;
                    string Result = cmd.ExecuteScalar().ToString();
                    conn.Close();

                    if (Result.Trim().ToLower() == "deleted")
                    {
                        //do not delete images as these might be there in another version for the same noun-modifier
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

                string FileName = "UserGuide-MRODictionary.pptx";

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
