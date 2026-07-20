using Aspose.Cells;
using Aspose.Words;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;
using GOPWebAPI.Models.Incident_Report_Models;
using Newtonsoft.Json.Linq;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Web;
using System.Web.Http;
using System.Xml.Linq;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/MarketingDocs")]
    public class MarketingDocsController : ApiController
    {
        private BLLMarketingDocs _BLLMarketingDocs;

        public MarketingDocsController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLMarketingDocs = new BLLMarketingDocs(connectionString);
        }

        #region Upload Marketing Document
        [HttpPost]
        [Route("UploadMarketingDocument")]
        public HttpResponseMessage UploadMarketingDocument([FromBody] MarketingDocModel model)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(model.UserID, "View Marketing Documents"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                //Check Model State
                if (!ModelState.IsValid)
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Invalid Data");

                string Result = _BLLMarketingDocs.UploadMarketingDocument(model);

                if (!string.IsNullOrEmpty(Result) && Result.Trim().ToLower().StartsWith("success"))
                {
                    string Domain = model.Domain;

                    if (model.Domain == "E")
                        Domain = "Engineering";
                    else if(model.Domain == "H")
                        Domain = "Healthcare";

                    DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));

                    if (File.Exists(dirTemp + model.FileName))
                    {
                        DirectoryInfo dirUploads = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/MarketingDocs/" + Domain + "/" + model.DocType + "/"));
                        FileOperations.MoveFile(dirTemp, model.FileName, dirUploads, model.UserUploadedFileName);
                    }

                    //return response status code
                    return Request.CreateResponse(HttpStatusCode.OK, "success");
                }
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Read All Marketing Documents from input domain and doc type
        [HttpGet]
        [Route("ReadAllMarketingDocuments")]
        public HttpResponseMessage ReadAllMarketingDocuments(string Domain, string DocType, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Marketing Documents"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                System.Data.DataTable dtDetails = new System.Data.DataTable();
                List<MarketingDocModel> lstMarketingDocs = new List<MarketingDocModel>();
                dtDetails = _BLLMarketingDocs.ReadAllMarketingDocuments(Domain, DocType);

                if (dtDetails.Rows.Count > 0)
                {
                    for (int i = 0; i < dtDetails.Rows.Count; i++)
                    {
                        MarketingDocModel marketingDocModel = new MarketingDocModel();
                        marketingDocModel.Id = Convert.ToInt32(dtDetails.Rows[i]["id"]);
                        marketingDocModel.FileName = dtDetails.Rows[i]["FileName"].ToString();
                        marketingDocModel.FileType = dtDetails.Rows[i]["FileType"].ToString();
                        marketingDocModel.UploadedOn = Convert.ToDateTime(dtDetails.Rows[i]["UploadedOn"]);
                        lstMarketingDocs.Add(marketingDocModel);
                    }

                    return Request.CreateResponse(HttpStatusCode.OK, lstMarketingDocs);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }

            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }

        #endregion

        #region Read Marketing Document Details by id
        [HttpGet]
        [Route("ReadMarketingDocumentById")]
        public HttpResponseMessage ReadMarketingDocumentById(int id)
        {
            try
            {
                System.Data.DataTable dtDetails = new System.Data.DataTable();
                MarketingDocModel marketingDocModel = new MarketingDocModel();
                dtDetails = _BLLMarketingDocs.ReadMarketingDocumentById(id);

                if (dtDetails.Rows.Count > 0)
                {
                    marketingDocModel.Id = Convert.ToInt32(dtDetails.Rows[0]["id"]);
                    marketingDocModel.FileName = dtDetails.Rows[0]["FileName"].ToString();
                    marketingDocModel.UserUploadedFileName = dtDetails.Rows[0]["FileName"].ToString();
                    marketingDocModel.Domain = dtDetails.Rows[0]["Domain"].ToString();
                    marketingDocModel.DocType = dtDetails.Rows[0]["DocType"].ToString();
                    marketingDocModel.FileType = dtDetails.Rows[0]["FileType"].ToString();
                    marketingDocModel.UploadedOn = Convert.ToDateTime(dtDetails.Rows[0]["UploadedOn"]);
                    marketingDocModel.UserID = dtDetails.Rows[0]["UploadedBy"].ToString();

                    return Request.CreateResponse(HttpStatusCode.OK, marketingDocModel);
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Msg = "No data found";
                    objResult.Success = 0;

                    return Request.CreateResponse(HttpStatusCode.OK, objResult);
                }

            }
            catch (Exception ex)
            {
                // Log the exception
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Download the file from input domain and doc. type
        [HttpGet]
        [Route("DownloadMarketingDocument")]
        public HttpResponseMessage DownloadMarketingDocument(string Domain, string DocType,string FileName)
        {
            try
            {
                if (Domain.Trim().ToUpper() == "E")
                    Domain = "Engineering";
                else if (Domain.Trim().ToUpper() == "H")
                    Domain = "Healthcare";

                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                //Create HTTP Response.
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                //Set the File Path.
                string filePath = HttpContext.Current.Server.MapPath("~/Uploads/MarketingDocs/" + Domain + "/" + DocType + "/") + FileName;

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

        #region Preview As Pdf
        [HttpGet]
        [Route("PreviewAsPdf")]
        public HttpResponseMessage PreviewAsPdf(string Domain, string DocType, string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "FileName is required.");

            try
            {
                // Build file path
                string basePath = HttpContext.Current.Server.MapPath("~/Uploads/MarketingDocs");
                string filePath = Path.Combine(basePath, Domain, DocType, fileName);

                if (!File.Exists(filePath))
                    return Request.CreateErrorResponse(HttpStatusCode.NotFound, "File not found.");

                string ext = Path.GetExtension(filePath).ToLower();
                MemoryStream pdfStream = new MemoryStream();

                // Convert according to extension
                switch (ext)
                {
                    case ".xlsx":
                    case ".xls":
                        Workbook wb = new Workbook();
                        Aspose.Cells.License l = new Aspose.Cells.License();
                        l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                        wb.Open(filePath);
                        wb.Save(pdfStream, Aspose.Cells.FileFormatType.Pdf);
                        break;

                    case ".docx":
                    case ".doc":
                        // Load the Word document
                        Spire.Doc.Document doc = new Spire.Doc.Document();
                        doc.LoadFromFile(filePath);

                        // Save to PDF in MemoryStream
                        pdfStream = new MemoryStream();
                        doc.SaveToStream(pdfStream, Spire.Doc.FileFormat.PDF);
                        break;
                    //Aspose.Words.License license = new Aspose.Words.License();
                    //license.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    //Document doc = new Document(filePath);
                    //doc.Save(pdfStream, Aspose.Words.SaveFormat.Pdf);
                    //break;

                    case ".pptx":
                    case ".ppt":
                        Presentation ppt = new Presentation();
                        ppt.LoadFromFile(filePath);

                        pdfStream = new MemoryStream();
                        ppt.SaveToFile(pdfStream, Spire.Presentation.FileFormat.PDF);
                        break;
                    //Aspose.Slides.License license1 = new Aspose.Slides.License();
                    //license1.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    //var ppt = new Aspose.Slides.Presentation(filePath);
                    //ppt.Save(pdfStream, Aspose.Slides.Export.SaveFormat.Pdf);
                    //break;

                    default:
                        return Request.CreateErrorResponse(
                            HttpStatusCode.BadRequest,
                            "Unsupported file type for preview."
                        );
                }

                pdfStream.Position = 0;

                // Prepare response as PDF
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new ByteArrayContent(pdfStream.ToArray());
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

                return response;
            }
            catch (Exception ex)
            {
                return Request.CreateErrorResponse(
                    HttpStatusCode.InternalServerError,
                    "Error converting file: " + ex.Message
                );
            }
        }
        #endregion

        #region Delete Marketing Document
        [HttpPost]
        [Route("DeleteMarketingDocument")]
        public HttpResponseMessage DeleteMarketingDocument(int id, string Domain, string DocType, string FileName, string UserID)
        {
            try
            {
                if (!AccessControl.CanUserAccessPage(UserID, "View Marketing Documents"))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Access Denied");

                string Result = _BLLMarketingDocs.DeleteMarketingDocument(id, UserID);

                if (!string.IsNullOrEmpty(Result) && Result.Trim().ToLower().StartsWith("success"))
                {
                    if (Domain == "E")
                        Domain = "Engineering";
                    else if (Domain == "H")
                        Domain = "Healthcare";

                    string UploadedFilePath = HttpContext.Current.Server.MapPath("~/Uploads/MarketingDocs/" + Domain + "/" + DocType + "/") + FileName;
                    if (File.Exists(UploadedFilePath))
                        File.Delete(UploadedFilePath);
                }

                return Request.CreateResponse(HttpStatusCode.OK);
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
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

                string FileName = "UserGuide-Marketing.pptx";

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
