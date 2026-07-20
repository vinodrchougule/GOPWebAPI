using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;

namespace GOPWebAPI.Controllers
{
    [RoutePrefix("api/ConfirmDuplicateSKUsFromCIF")]
    public class ConfirmDuplicateSKUsFromCIFController : ApiController
    {
        private BLLConfirmDuplicateSKUs _BLLConfirmDuplicateSKUs;
        public ConfirmDuplicateSKUsFromCIFController()
        {
            string connectionString = ConfigurationManager.ConnectionStrings["GOPDBConnectionString"].ConnectionString;
            _BLLConfirmDuplicateSKUs = new BLLConfirmDuplicateSKUs(connectionString);
        }

        #region Read Duplicate SKUs based on Selected Columns
        [HttpPost]
        [Route("ReadDuplicateSKUs")]
        public IHttpActionResult ReadDuplicateSKUs(DuplicateToFindOnColumnsModel model)
        {
            try
            {
                List<int> lstColumnNosOfSelectedColumns = new List<int>();
                int DuplicateRemarksColNo = -1, DuplicateRowSetCounter = 1, TotalDataRows = 0, DuplicateRowsCount = 0;
                bool cellValueMatched = false, AreAllColumnValuesBlank = false, IsDuplicateCounterPrinted = false;
                DataFormatConverter dataFormatConverter = new DataFormatConverter();

                #region Check if SQL Table already exists
                string Result = _BLLConfirmDuplicateSKUs.IsCIFSQLTableExists(model.CustomerCode, model.ProjectCode, model.BatchNo);
                #endregion

                if (Result.Trim().ToLower() == "no")
                {
                    #region Find the customer input file name
                    string FileName = string.Empty;
                    DirectoryInfo dirCustomerInputFile;

                    if (string.IsNullOrEmpty(model.BatchNo))
                        FileName = model.CustomerCode + '_' + model.ProjectCode + "_CustomerInputFile.xlsx";
                    else
                        FileName = model.CustomerCode + '_' + model.ProjectCode + '_' + model.BatchNo + "_CustomerInputFile.xlsx";

                    if (string.IsNullOrEmpty(model.BatchNo))
                        dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                    else
                        dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));

                    string CustomerInputFilepath = dirCustomerInputFile + FileName;

                    if (!File.Exists(CustomerInputFilepath))
                        return Content(HttpStatusCode.BadRequest, "Customer Input File Not Found");
                    #endregion

                    #region Copy the file to temp folder with unique name
                    DirectoryInfo dirTemp = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/temp/"));
                    string TempCustomerFileName = Guid.NewGuid().ToString() + ".xlsx";
                    string TempCustomerFileFullPath = dirTemp.FullName + TempCustomerFileName;
                    FileInfo fileInfo = new FileInfo(CustomerInputFilepath);
                    fileInfo.CopyTo(TempCustomerFileFullPath);
                    #endregion

                    #region Open Temp Customer Input File and add Duplicate Remarks column
                    Workbook wbTIF = new Workbook();             //Input File
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    wbTIF.LoadData(TempCustomerFileFullPath);
                    var wsTIF = wbTIF.Worksheets[0];
                    TotalDataRows = wsTIF.Cells.MaxRow;

                    DuplicateRemarksColNo = wsTIF.Cells.MaxColumn + 1;
                    wsTIF.Cells[0, DuplicateRemarksColNo].PutValue("DuplicateRemarks");
                    #endregion

                    #region Find the Column Nos. of Selected Columns
                    foreach (DuplicateToFindOnColumns columnName in model.ColumnNames)
                    {
                        for (int col = 0; col <= wsTIF.Cells.MaxColumn; col++)
                        {
                            if (wsTIF.Cells[0, col].StringValue.Trim().ToUpper() == columnName.ColumnName.Trim().ToUpper())
                            {
                                lstColumnNosOfSelectedColumns.Add(col);
                                break;
                            }
                        }
                    }
                    #endregion

                    #region Find the duplicate SKUs and write the row set no. in Duplicate Remarks column
                    for (int mainRow = 1; mainRow <= wsTIF.Cells.MaxRow; mainRow++)
                    {
                        if (IsDuplicateCounterPrinted)
                        {
                            DuplicateRowSetCounter++;
                            IsDuplicateCounterPrinted = false;
                        }

                        if (string.IsNullOrEmpty(wsTIF.Cells[mainRow, DuplicateRemarksColNo].StringValue))
                        {
                            AreAllColumnValuesBlank = false;
                            foreach (int c in lstColumnNosOfSelectedColumns)
                            {
                                if (!string.IsNullOrEmpty(wsTIF.Cells[mainRow, c].StringValue))
                                    break;
                                else
                                    AreAllColumnValuesBlank = true;
                            }

                            if (!AreAllColumnValuesBlank)
                            {
                                for (int iRow = mainRow + 1; iRow <= wsTIF.Cells.MaxRow; iRow++)
                                {
                                    if (string.IsNullOrEmpty(wsTIF.Cells[iRow, DuplicateRemarksColNo].StringValue))
                                    {
                                        cellValueMatched = false;
                                        foreach (int col in lstColumnNosOfSelectedColumns)
                                        {
                                            if (dataFormatConverter.RemoveSpecialCharacters(wsTIF.Cells[mainRow, col].StringValue.Trim().ToUpper()) == dataFormatConverter.RemoveSpecialCharacters(wsTIF.Cells[iRow, col].StringValue.Trim().ToUpper()))
                                                cellValueMatched = true;
                                            else
                                            {
                                                cellValueMatched = false;
                                                break;
                                            }
                                        }

                                        if (cellValueMatched)
                                        {
                                            if (string.IsNullOrEmpty(wsTIF.Cells[mainRow, DuplicateRemarksColNo].StringValue))
                                                wsTIF.Cells[mainRow, DuplicateRemarksColNo].PutValue("Duplicate Row Set " + DuplicateRowSetCounter.ToString());
                                            wsTIF.Cells[iRow, DuplicateRemarksColNo].PutValue("Duplicate Row Set " + DuplicateRowSetCounter.ToString());
                                            IsDuplicateCounterPrinted = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region Add ID column and update row no.
                    int IDColNo = DuplicateRemarksColNo + 1;
                    wsTIF.Cells[0, IDColNo].PutValue("ID");
                    for (int row = 1; row <= wsTIF.Cells.MaxRow; row++)
                        wsTIF.Cells[row, IDColNo].PutValue(row);
                    #endregion

                    #region Delete Duplicate Remarks blank rows from the temp Customer Input File
                    for (int row = 1; row <= wsTIF.Cells.MaxRow;)
                    {
                        if (string.IsNullOrEmpty(wsTIF.Cells[row, DuplicateRemarksColNo].StringValue.Trim()))
                        {
                            wsTIF.Cells.DeleteRow(row);
                            if (wsTIF.Cells.MaxRow == 0 || row == wsTIF.Cells.MaxRow)
                                break;
                        }
                        else
                            row++;
                    }

                    if (string.IsNullOrEmpty(wsTIF.Cells[wsTIF.Cells.MaxRow, DuplicateRemarksColNo].StringValue.Trim()))
                        wsTIF.Cells.DeleteRow(wsTIF.Cells.MaxRow);
                    #endregion

                    #region Write Duplicate Percentage
                    DuplicateRowsCount = wsTIF.Cells.MaxRow;
                    wsTIF.Cells[0, IDColNo + 1].PutValue("Duplicate Percentage");
                    wsTIF.Cells[1, IDColNo + 1].PutValue((DuplicateRowsCount * 100.00) / TotalDataRows);
                    #endregion

                    #region Save the file changes and convert the file to datatable
                    wbTIF.Save(TempCustomerFileFullPath);
                    DataTable excelData = dataFormatConverter.ExcelToDataTable(TempCustomerFileFullPath);

                    string ColumnsToBeSorted = string.Empty;
                    foreach (DuplicateToFindOnColumns colName in model.ColumnNames)
                    {
                        if (string.IsNullOrEmpty(ColumnsToBeSorted))
                            ColumnsToBeSorted = colName.ColumnName + " ASC";
                        else
                            ColumnsToBeSorted += "," + colName.ColumnName + " ASC";
                    }
                    #endregion

                    string selectedColumnNames1 = string.Empty;
                    foreach (DuplicateToFindOnColumns colName in model.ColumnNames)
                    {
                        if (string.IsNullOrEmpty(selectedColumnNames1))
                            selectedColumnNames1 = "[" + colName.ColumnName + "]";
                        else
                            selectedColumnNames1 += "," + "[" + colName.ColumnName + "]";
                    }

                    DataView dv = excelData.DefaultView;
                    dv.Sort = ColumnsToBeSorted;
                    DataTable sortedDT = dv.ToTable();

                    DataTable orderedTable = new DataTable();
                    orderedTable.Columns.Add("ID");
                    orderedTable.Columns.Add("DuplicateRemarks");
                    orderedTable.Columns.Add("SelectedColumns");
                    orderedTable.Columns.Add("IsDuplicate", typeof(bool));

                    foreach (DataColumn column in sortedDT.Columns)
                    {
                        if (column.ColumnName != "ID" && column.ColumnName != "SelectedColumns" && column.ColumnName != "DuplicateRemarks")
                        {
                            orderedTable.Columns.Add(column.ColumnName, column.DataType);
                        }
                    }

                    foreach (DataRow row in sortedDT.Rows)
                    {
                        orderedTable.ImportRow(row);
                    }

                    foreach (DataRow row in orderedTable.Rows)
                    {
                        row["IsDuplicate"] = true;
                    }

                    if(model.ColumnNames.Count > 0)
                    {
                        foreach (DataRow row in orderedTable.Rows)
                        {
                            row["SelectedColumns"] = selectedColumnNames1;
                        }
                    }

                    #region Delete the temp file
                    FileInfo TempCustomerFileInfo = new FileInfo(TempCustomerFileFullPath);
                    TempCustomerFileInfo.Delete();
                    #endregion

                    orderedTable.Columns["ID"].SetOrdinal(0);
                    orderedTable.Columns["DuplicateRemarks"].SetOrdinal(1);
                    int ordinal = 2;
                    foreach (DataColumn column in sortedDT.Columns)
                    {
                        if (column.ColumnName != "ID" && column.ColumnName != "SelectedColumns" && column.ColumnName != "DuplicateRemarks")
                        {
                            foreach (DuplicateToFindOnColumns colName in model.ColumnNames)
                            {
                                if(column.ColumnName.Trim().ToLower() == colName.ColumnName.Trim().ToLower())
                                {
                                    orderedTable.Columns[column.ColumnName].SetOrdinal(ordinal);
                                    ordinal++;
                                }
                            }
                        }
                    }
                    orderedTable.Columns["SelectedColumns"].SetOrdinal(ordinal);
                    ordinal++;
                    orderedTable.Columns["IsDuplicate"].SetOrdinal(ordinal);

                    //return Ok(sortedDT);
                    return Ok(orderedTable); 
                }
                else
                {
                    #region Fetch CIF SKUs From SQL Table
                    string selectedColumnNames = string.Empty;
                    foreach(DuplicateToFindOnColumns colName in model.ColumnNames)
                    {
                        if (string.IsNullOrEmpty(selectedColumnNames))
                            selectedColumnNames = "[" + colName.ColumnName + "]";
                        else
                            selectedColumnNames += "," + "[" + colName.ColumnName + "]";
                    }

                    ConfirmDuplicateSKUsModel confirmDuplicateSKUsModel = new ConfirmDuplicateSKUsModel();
                    confirmDuplicateSKUsModel.CustomerCode = model.CustomerCode;
                    confirmDuplicateSKUsModel.ProjectCode = model.ProjectCode;
                    confirmDuplicateSKUsModel.BatchNo = model.BatchNo;
                    confirmDuplicateSKUsModel.SelectedColumns = selectedColumnNames;

                    DataTable dataTable = _BLLConfirmDuplicateSKUs.FetchCIFSKUsFromTable(confirmDuplicateSKUsModel);

                    string ColumnsToBeSorted = string.Empty;
                    foreach (DuplicateToFindOnColumns colName in model.ColumnNames)
                    {
                        if (string.IsNullOrEmpty(ColumnsToBeSorted))
                            ColumnsToBeSorted = "[" + colName.ColumnName + "] ASC";
                        else
                            ColumnsToBeSorted += ",[" + colName.ColumnName + "] ASC";
                    }

                    DataView dv = dataTable.DefaultView;
                    dv.Sort = ColumnsToBeSorted;
                    DataTable sortedDT = dv.ToTable();
                    #endregion

                    return Ok(sortedDT);
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

        #region Write CIF Data To Database
        [HttpPost]
        [Route("WriteCIFDataToDatabase")]
        public HttpResponseMessage WriteCIFDataToDatabase(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            try
            {
                string sqlTableName = string.Empty, FileName = string.Empty;
                DirectoryInfo dirCustomerInputFile;

                if (string.IsNullOrEmpty(BatchNo))
                    BatchNo = "";

                #region Get Customer Input File path and SQL Table Name
                if (string.IsNullOrEmpty(BatchNo))
                    FileName = CustomerCode + '_' + ProjectCode + "_CustomerInputFile.xlsx";
                else
                    FileName = CustomerCode + '_' + ProjectCode + '_' + BatchNo + "_CustomerInputFile.xlsx";

                if (string.IsNullOrEmpty(BatchNo))
                    sqlTableName = "CustomerInputFile_" + CustomerCode + '_' + ProjectCode;
                else
                    sqlTableName = "CustomerInputFile_" + CustomerCode + '_' + ProjectCode + '_' + BatchNo;

                if (string.IsNullOrEmpty(BatchNo))
                    dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/"));
                else
                    dirCustomerInputFile = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/Uploads/GPMT/CustomerInputFile/Batch/"));

                string CustomerInputFilepath = dirCustomerInputFile + FileName;

                if (!File.Exists(CustomerInputFilepath))
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Customer Input File Not Found");
                #endregion

                #region Check if SQL Table already exists
                string Result = _BLLConfirmDuplicateSKUs.IsCIFSQLTableExists(CustomerCode, ProjectCode, BatchNo);
                #endregion

                if (Result.Trim().ToLower() == "no")
                {
                    #region Form string to create table
                    #region Open the File as Work Book and get first worksheet
                    Workbook wbCIF = new Workbook();                       //Customer Input File Work Book
                    Aspose.Cells.License l = new Aspose.Cells.License();
                    l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                    wbCIF.LoadData(CustomerInputFilepath);
                    Worksheet wsCIF = wbCIF.Worksheets[0];
                    #endregion

                    #region Generate dynamic column name and data type string
                    string ColumnNameDataTypeString = string.Empty, ColumnName = string.Empty, ColumnDataType = string.Empty, CellValue = string.Empty;
                    string strSQL = "CREATE TABLE " + sqlTableName + "(ID INT IDENTITY(1,1),";
                    for (int col = 0; col <= wsCIF.Cells.MaxColumn; col++)
                    {
                        if (!string.IsNullOrEmpty(wsCIF.Cells[0, col].StringValue.Trim()))
                        {
                            ColumnName = "[" + wsCIF.Cells[0, col].StringValue.Trim() + "]";
                            ColumnDataType = "VARCHAR(4000)";

                            if (string.IsNullOrEmpty(ColumnNameDataTypeString.Trim()))
                                ColumnNameDataTypeString = ColumnName + " " + ColumnDataType;
                            else
                                ColumnNameDataTypeString += "," + ColumnName + " " + ColumnDataType;
                        }
                    }
                    #endregion

                    strSQL += ColumnNameDataTypeString + ",IsDuplicate BIT DEFAULT(0))";
                    #endregion

                    #region Create CIF SQL Table in database
                    Result = _BLLConfirmDuplicateSKUs.CreateCIFSQLTable(CustomerCode, ProjectCode, BatchNo, strSQL);
                    #endregion

                    if (Result.StartsWith("Success"))
                    {
                        #region Writing data to SQL table
                        SqlConnection conn = new SqlConnection(DBConnInfo.ConnectionString());
                        DataFormatConverter dataFormatConverter = new DataFormatConverter();
                        DataTable excelData = dataFormatConverter.ExcelToDataTable(CustomerInputFilepath);
                        SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                        bulkcopy.DestinationTableName = sqlTableName;
                        foreach (DataColumn col in excelData.Columns)
                        {
                            bulkcopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                        }
                        conn.Open();
                        bulkcopy.WriteToServer(excelData);
                        conn.Close();
                        #endregion
                    }
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Customer Input File data updated to database successfully");
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Update selected SKUs as Duplicates in CIF table on selected columns
        [HttpPost]
        [Route("UpdateSelectedSKUsAsDuplicates")]
        public HttpResponseMessage UpdateSelectedSKUsAsDuplicates([FromBody] ConfirmDuplicateSKUsModel model)
        {
            try
            {
                string Result = _BLLConfirmDuplicateSKUs.UpdateSelectedSKUsAsDuplicates(model);
                if (Result.StartsWith("Success"))
                    return Request.CreateResponse(HttpStatusCode.OK, "Selected SKUs marked as Duplicates successfully");
                else
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, Result);
            }
            catch (Exception ex)
            {
                //log error to database table and return error response
                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Check if duplicate SKUs already saved in database
        [HttpGet]
        [Route("AreDuplicateSKUsAlreadySaved")]
        public HttpResponseMessage AreDuplicateSKUsAlreadySaved(string CustomerCode, string ProjectCode, string BatchNo = "")
        {
            string Result = _BLLConfirmDuplicateSKUs.IsCIFSQLTableExists(CustomerCode, ProjectCode, BatchNo);
            return Request.CreateResponse(HttpStatusCode.OK, Result);
        }
        #endregion

    }
}
