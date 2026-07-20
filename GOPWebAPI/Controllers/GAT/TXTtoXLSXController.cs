using Aspose.Cells;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using Microsoft.Ajax.Utilities;
using Swashbuckle.Swagger;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;
using System.Web.Http.Results;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/TXTtoXLSX")]
    public class TXTtoXLSXController : ApiController
    {
        #region Convert TXT to XLSX input uploaded file
        [HttpPost]
        [Route("ConvertTXTtoXLSX")]
        public HttpResponseMessage ConvertTXTtoXLSX([FromBody] ConversionRequest conversionRequest)
        {
            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + conversionRequest.InputFileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                Workbook workbook = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                workbook.Open(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
                var sheet = workbook.Worksheets[0];

                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(conversionRequest.UploadedInputFileName);
                string OutputFileName = "Output_" + fileNameWithoutExtension + ".xlsx";

                // Read raw bytes to preserve embedded line breaks inside field values
                string rawContent = File.ReadAllText(InputFilepath);

                // Normalize line endings: \r\n and \r → \n
                rawContent = rawContent.Replace("\r\n", "\n").Replace("\r", "\n");

                // Split into lines (now all endings are \n)
                string[] lines = rawContent.Split('\n');

                // Remove completely empty trailing lines caused by final newline in file
                lines = lines.Where(l1 => !string.IsNullOrEmpty(l1)).ToArray();

                if (lines.Length == 0)
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input file is empty.");
                }

                string[] headerParts = lines[0].Split(
                    new string[] { conversionRequest.Delimiter }, StringSplitOptions.None);
                int expectedColumnCount = headerParts.Length;
                int excelRowIndex = 0;

                for (int i = 0; i < lines.Length; i++)
                {
                    string currentLine = lines[i];

                    // Skip blank lines
                    if (string.IsNullOrWhiteSpace(currentLine))
                        continue;

                    string[] parts = currentLine.Split(
                        new string[] { conversionRequest.Delimiter }, StringSplitOptions.None);

                    // Merge subsequent lines if column count is less than expected.
                    // Fewer columns = unexpected line break inside a field value.
                    while (parts.Length < expectedColumnCount && (i + 1) < lines.Length)
                    {
                        i++;
                        currentLine += lines[i];
                        parts = currentLine.Split(
                            new string[] { conversionRequest.Delimiter }, StringSplitOptions.None);
                    }

                    // Validate column count AFTER merging all broken lines
                    //if (parts.Length > expectedColumnCount)
                    //{
                    //    if (File.Exists(InputFilepath))
                    //        File.Delete(InputFilepath);

                    //    Result objResult1 = new Result();
                    //    objResult1.Success = 0;
                    //    objResult1.Msg = $"Line No. {i + 1} has {parts.Length} values " +
                    //                     $"but only {expectedColumnCount} column headers exist.";
                    //    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult1.Msg);
                    //}

                    // Write columns to Excel row
                    for (int j = 0; j < parts.Length; j++)
                    {
                        string value = conversionRequest.IsToTrimSplittedValues
                            ? parts[j].Trim()
                            : parts[j];

                        if (value.Length > 0)
                            sheet.Cells[excelRowIndex, j].PutValue(value);
                    }

                    excelRowIndex++;
                }

                #region Add Observation column at the end
                int LastColumnNo = sheet.Cells.MaxColumn;
                int ObservationColumnNo = LastColumnNo + 1;
                sheet.Cells[0, ObservationColumnNo].PutValue("Observation");

                int TotalRows = sheet.Cells.MaxRow;
                string[] formats = {
                                    "M/d/yyyy", "MM/dd/yyyy", "d/M/yyyy", "dd/MM/yyyy",
                                    "yyyy-MM-dd", "yyyy/MM/dd", "dd-MMM-yyyy", "MMM dd, yyyy",
                                    "dd.MM.yyyy", "yyyy.MM.dd" // Added dots as separators
                                    };

                for (int col = 0; col <= LastColumnNo; col++)
                {
                    int intRows = 0, dateRows = 0, decimalRows = 0, stringRows = 0;
                    string columnDataType = "string";
                    string columnName = sheet.Cells[0, col].StringValue;

                    // Step 1: Detect Majority Data Type
                    for (int row = 1; row <= TotalRows; row++)
                    {
                        string cellValue = sheet.Cells[row, col].StringValue?.Trim();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        if (int.TryParse(cellValue, out _))
                            intRows++;
                        else if (decimal.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                            decimalRows++;
                        else if (DateTime.TryParseExact(cellValue, formats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out _))
                            dateRows++;
                        else
                            stringRows++;
                    }

                    // Step 2: Determine Column Type Based on Majority
                    // Logic: If decimals exist, it's a decimal column even if many are ints.
                    decimal totalPopulated = intRows + dateRows + decimalRows + stringRows;
                    if (totalPopulated == 0) continue; // Skip empty columns

                    if (dateRows > intRows && dateRows > decimalRows && dateRows > stringRows)
                        columnDataType = "date";
                    else if (stringRows > intRows && stringRows > dateRows && stringRows > decimalRows)
                        columnDataType = "string";
                    else if (decimalRows > 0 || intRows > 0)
                    {
                        // If there are any decimals, treat the whole column as decimal
                        columnDataType = decimalRows > 0 ? "decimal" : "int";
                    }

                    // Step 3: Validate Rows against determined Column Type
                    for (int row = 1; row <= TotalRows; row++)
                    {
                        string cellValue = sheet.Cells[row, col].StringValue?.Trim();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        bool isValid = true;
                        string errorMessage = "";

                        switch (columnDataType)
                        {
                            case "int":
                                isValid = int.TryParse(cellValue, out _);
                                errorMessage = $"Not an integer value in column '{columnName}'";
                                break;
                            case "date":
                                isValid = DateTime.TryParseExact(cellValue, formats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out _);
                                errorMessage = $"Not a date value in column '{columnName}'";
                                break;
                            case "decimal":
                                isValid = decimal.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out _);
                                errorMessage = $"Not a decimal value in column '{columnName}'";
                                break;
                        }

                        if (!isValid)
                        {
                            Cell obsCell = sheet.Cells[row, ObservationColumnNo];
                            string existingObs = obsCell.StringValue;

                            if (string.IsNullOrEmpty(existingObs))
                                obsCell.PutValue(errorMessage);
                            else if (!existingObs.Contains(errorMessage)) // Prevent duplicate messages
                                obsCell.PutValue(existingObs + ", " + errorMessage);
                        }
                    }
                }
                #endregion

                bool IsObservationExists = false;
                for (int row = 1; row <= sheet.Cells.MaxRow; row++)
                {
                    if (!string.IsNullOrEmpty(sheet.Cells[row, ObservationColumnNo].StringValue.Trim()))
                    {
                        IsObservationExists = true;
                        break;
                    }
                }

                if (!IsObservationExists)
                    sheet.Cells.DeleteColumn(ObservationColumnNo);

                #region Saving and downloading the report
                sheet.AutoFitColumns();
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + OutputFileName;
                workbook.Save(filename);

                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
                #endregion
            }
            catch (Exception ex)
            {
                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Convert the splitted data to common template
        [HttpPost]
        [Route("ConvertSplittedDataToCommonTemplate")]
        public HttpResponseMessage ConvertSplittedDataToCommonTemplate(string UploadedInputFileName, string InputFileName, string HospitalName, string ItemOrPO)
        {
            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
            string InValidColumnName = string.Empty;

            try
            {
                Workbook workbook = new Workbook();

                if(HospitalName.Trim().ToLower() == "albany" || HospitalName.Trim().ToLower() == "mercy" ||
                   HospitalName.Trim().ToLower() == "childrensconnecticut" || HospitalName.Trim().ToLower() == "minnesota" ||
                   HospitalName.Trim().ToLower() == "nuvance" || HospitalName.Trim().ToLower() == "texas" ||
                   HospitalName.Trim().ToLower() == "uhs" || HospitalName.Trim().ToLower() == "centracare")
                {
                    InValidColumnName = IsValidData(InputFileName, HospitalName.ToLower(), ItemOrPO);
                    if (string.IsNullOrEmpty(InValidColumnName))
                    {
                        workbook = ConvertAndGetTheWorkbook(InputFileName, HospitalName.ToLower(), ItemOrPO);
                        Worksheet ws = workbook.Worksheets[0];
                        Cells cells = ws.Cells;

                        int totalRows = cells.MaxDataRow + 1;
                        int totalCols = cells.MaxDataColumn + 1;

                        HashSet<string> seenRows = new HashSet<string>();
                        List<int> rowsToDelete = new List<int>();

                        for (int i = 0; i < totalRows; i++)
                        {
                            System.Text.StringBuilder rowKey = new System.Text.StringBuilder();
                            for (int j = 0; j < totalCols; j++)
                            {
                                Cell cell = cells[i, j];
                                rowKey.Append(cell == null ? "" : cell.StringValue);
                                rowKey.Append("|");  // delimiter to avoid false matches
                            }

                            string key = rowKey.ToString();
                            if (seenRows.Contains(key))
                                rowsToDelete.Add(i);
                            else
                                seenRows.Add(key);
                        }

                        // Delete from bottom up to avoid index shifting
                        for (int i = rowsToDelete.Count - 1; i >= 0; i--)
                        {
                            cells.DeleteRow(rowsToDelete[i]);
                        }

                        workbook.Save(InputFilepath);
                    }
                    else
                    {
                        Result objResult1 = new Result();
                        objResult1.Success = 0;
                        objResult1.Msg = "Column '" + InValidColumnName + "' not found in input item file data for the hospital - " + HospitalName;
                        return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult1.Msg);
                    }
                }
                else
                {
                    Result objResult = new Result();
                    objResult.Success = 0;
                    objResult.Msg = "No conversion logic defined for the hospital - " + HospitalName;
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
                }

                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(UploadedInputFileName);
                //string OutputFileName = "Output_In_Template_Format_" + fileNameWithoutExtension + "_" + HospitalName +".xlsx";
                string OutputFileName = fileNameWithoutExtension + ".xlsx";
                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + OutputFileName;
                workbook.Save(filename);

                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
            }
            catch (Exception ex)
            {
                if (File.Exists(InputFilepath))
                    File.Delete(InputFilepath);

                ExceptionLogging.SendExceptionToDB(ex);
                Result objResult = new Result();
                objResult.Success = 0;
                objResult.Msg = ex.Message;
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, objResult.Msg);
            }
        }
        #endregion

        #region Validate input file data
        private string IsValidData(string InputFileName, string HospitalName, string ItemOrPO)
        {
            string InvalidColumnName = string.Empty;
            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

            Aspose.Cells.License l = new Aspose.Cells.License();
            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));

            Workbook workbook = new Workbook();
            workbook.Open(InputFilepath);
            Worksheet ws = workbook.Worksheets[0];

            string[] inputColumns = null;
            string mode = ItemOrPO.Trim().ToLower();

            #region Albany
            if (HospitalName.Trim().ToLower() == "albany")
            {

                if (mode == "item")
                {
                    inputColumns = new[] {
                                                "Organiztion", "HospitalName", "ItemNumber", "ItemDescription",
                                                "VendorName", "VendorCatalogNumber", "MFRName", "MFRCatalogNumber",
                                                "UOM", "Conversion", "UOMPrice", "StockStatus", "StockUOM",
                                                "IssueUOM", "UNSPSCCode", "UNSPSCCategory", "ContractNumber",
                                                "ContractType", "ContractDescription"
                                          };
                }
                else if (mode == "po")
                {
                    inputColumns = new[] {
                                            "Organization", "HospitalName", "PONumber", "POLineNumber",
                                            "PODate", "DepartmentName", "ItemNumber", "ItemDescription",
                                            "VendorName", "VendorCatalogNumber", "MFRName", "MFRCatalogNumber",
                                            "UOM", "Conversion", "UOMPrice", "POQty", "POLineSpend"
                                          };
                }
            }
            #endregion
            #region Mercy
            else if (HospitalName.Trim().ToLower() == "mercy")
            {
                if (mode == "item")
                {
                    inputColumns = new[] {
                                            "Facility","Hospital_Item_Number", "Item_Description",
                                            "Purchase Unit of Measure","Conversion Factor","Stock_UOM","Vendor Name", "Vendor Catalog Number",
                                            "Manufacturer Name", "Manufacturer Catalog Number","Status","Price Per UOM",
                                            "Contract Number","Commodity Description"
                                         };
                }
                else if (mode == "po")
                {
                    inputColumns = new[] {
                                            "PO_Date","PO_Number", "Line_Nbr",
                                            "Facility","Department Name","Item",
                                            "Vendor Name", "Ven_Item","Manufacturer Name", "Manufacturer Product Number",
                                            "Description","Conversion Factor","Quantity","PUOM",
                                            "Price Per UOM", "PO Line Spend"
                                         };
                }
            }
            #endregion
            #region Children's Connecticut
            else if (HospitalName.Trim().ToLower() == "childrensconnecticut")
            {
                if (mode == "item")
                {
                    inputColumns = new[] {
                                            "Item Number","Item Description",
                                            "Manufacturer Name","Manufacturer Catalog Number","Vendor Name","Vendor Catalog Number",
                                            "Purchase UOM","Purchase UOM Pack Factor","Current Price","Item Type","UNSPSC Code", "UNSPSC Description",
                                            "Stock UOM","Issue UOM"
                                         };
                }
                else if (mode == "po")
                {
                    inputColumns = new[] {
                                            "Item Number","Item Description",
                                            "Manufacturer Name","Manufacturer Catalog Number","Vendor Name","Vendor Catalog Number",
                                            "Department Name","PO Date","PO Number","PO Line Number",
                                            "Purchase UOM","Quantity per UOM (pack factor)","Quantity Purchased","Price Paid (PO Price)","Extended Price Paid (PO Price)"
                                         };
                }
            }
            #endregion
            #region Minnesota
            else if (HospitalName.Trim().ToLower() == "minnesota")
            {
                if (mode == "item")
                {
                    inputColumns = new[] {
                                            "Hospital Item Number","Item Description",
                                            "Purchase UOM","Conversion factor","Vendor Name","Vendor Catalog Number",
                                            "Manufacturer Name","Manufacturer Catalog Number","Status","Price Per UOM",
                                            "Contract Number","UNSPSC Commodity Code", "UNSPSC Commodity Description"
                                         };
                }
                else if (mode == "po")
                {

                    inputColumns = new[] {
                                            "PO Date","PO Number",
                                            "PO Line Number","Department Name",
                                            "Hospital Item Number","Vendor Name","Vendor Product Number",
                                            "Manufacturer Name","Manufacturer's Product Number","Item Description","Conversion Factor",
                                            "PO Quantity","PO UOM","PO Price per UOM","PO Line Spend"
                                         };
                }
            }
            #endregion
            #region nuvance
            else if (HospitalName.Trim().ToLower() == "nuvance")
            {
                if (mode == "item")
                {
                    inputColumns = new[] {
                                            "Hospital Item Number","Hospital Item Description",
                                            "Manufacturer Name","Manufacturer Catalog Number","Vendor Name","Vendor Catalog Number",
                                            "Buy UOM","Buy UOM Conversion","Current Price","UNSPSC Commodity Code", "UNSPSC Commodity Description"
                                         };
                }
                else if(mode == "po")
                {
                    inputColumns = new[] {
                                            "Hospital Name","Hospital Item Number","Hospital Item Description",
                                            "Manufacturer Name","Manufacturer Catalog Number","Hospital Vendor Name","Vendor Catalog Number",
                                            "Hospital Department Name","PO Date","PO Number","PO Line Number",
                                            "Purchase UOM","Quantity per UOM","Quantity Purchased","Price Paid (PO Price)","Extended Price Paid (PO Price)"
                                         };
                }
            }
            #endregion
            #region texas
            else if (HospitalName.Trim().ToLower() == "texas")
            {
                if (mode == "item")
                {
                    inputColumns = new[] {
                                            "INV_ITEM_ID","DESCR254_MIXED",
                                            "UNIT_OF_MEASURE","CONVERSION_RATE","NAME1","ITM_ID_VNDR",
                                            "DESCR60","MFG_ITM_ID","STOCK_SYMBOL","PRICE_VNDR",
                                            "MODEL","CATEGORY_CD","DESCR60_2"
                                         };
                }
                else if(mode == "po")
                {
                    inputColumns = new[] {
                                            "PO_DT","PO_ID","LINE_NBR",
                                            "DESCR","INV_ITEM_ID","CONVERSION_RATE",
                                            "NAME1","ITM_ID_VNDR","DESCR60","MFG_ITM_ID",
                                            "DESCR254_MIXED","QTY_PO","UNIT_OF_MEASURE","PRICE_PO","TCH_SPENDING_AMT"
                                         };
                }
            }
            #endregion
            #region uhs
            else if (HospitalName.Trim().ToLower() == "uhs")
            {
                if (mode == "item")
                {
                    inputColumns = new[] {
                                            "HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION", "MANUFACTURER_NAME","MANUFACTURER_CATALOG_NUMBER",
                                            "VENDOR_NAME","VENDOR_CATALOG_NUMBER","PACK_UOM","PURCHASE_UOM_PACK_FACTOR","CURRENT_PRICE","ITEM_TYPE",
                                            "UNSPSC_COMMODITY_CODE","UNSPSC_COMMODITY_DESCRIPTION","Stock UOM","Issue UOM"
                                         };
                }
                else if (mode == "po")
                {
                    inputColumns = new[] {
                                            "HOSPITAL_NAME","HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION","MANUFACTURER_NAME","MANUFACTURER_CATALOG_NUMBER",
                                            "VENDOR_NAME","VENDOR_CATALOG_NUMBER","HOSPITAL_DEPARTMENT_NAME","PO_DATE","PO_NUMBER","PO_LINE_NUMBER",
                                            "PURCHASE_UOM","QUANTITY_PER_UOM","QUANTITY_PURCHASED","PO_PRICE_PAID","PO_EXTENDED_PRICE_PAID"
                                         };
                }
            }
            #endregion
            #region Centra Care
            else if (HospitalName.Trim().ToLower() == "centracare")
            {
                if (mode == "item")
                {
                    inputColumns = new[] {
                                            "HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION",
                                            "MANUFACTURER_NAME","MANUFACTURER_CATALOG_NUMBER", "VENDOR_NAME","VENDOR_CATALOG_NUMBER", "ORDER_UOM",
                                            "ORDER_UOM_MULT", "ORDER_PRICE","ITEM_TYPE", "AGREEMENT",
                                            "UNSPSC_CODE","UNSPSC_DESCRIPTION"
                                         };
                }
                else if (mode == "po")
                {
                    inputColumns = new[] {
                                            "FACILITY_NAME","CREATE_DATE","PO_NUMBER",
                                            "PO_LINE_NUM",
                                            "HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION","MANUFACTURER",
                                            "MANUFACTURER_PART_NUM","VENDOR_NAME","VENDOR_PRODUCT_NUM",
                                            "ORDER_UOM","ORDER_UOM_MULT",
                                            "ORDER_QUANTITY","ORDER_UNIT_PRICE","ORDER_EXTENDED_PRICE",
                                            "DEPARTMENT_DESCRIPTION"
                                         };
                }
            }
            #endregion

            if (inputColumns != null)
            {
                foreach (string requiredCol in inputColumns)
                {
                    bool isColumnFound = false;

                    for (int col = 0; col <= ws.Cells.MaxColumn; col++)
                    {
                        if (string.Equals(ws.Cells[0, col].StringValue.Trim(), requiredCol, StringComparison.OrdinalIgnoreCase))
                        {
                            isColumnFound = true;
                            break;
                        }
                    }

                    if (!isColumnFound)
                    {
                        InvalidColumnName = requiredCol;
                        break;
                    }
                }
            }
            
            return InvalidColumnName;
        }
        #endregion

        #region Convert And Get The Workbook
        private Workbook ConvertAndGetTheWorkbook(string InputFileName, string HospitalName, string ItemOrPO)
        {
            string InputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);
            Workbook workbook = new Workbook();
            Aspose.Cells.License l = new Aspose.Cells.License();
            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
            workbook.Open(InputFilepath);

            string mode = ItemOrPO.Trim().ToLower();

            string[] outputColumns = null, inputColumns = null;

            Workbook outputWorkbook = new Workbook();

            if (mode == "item")
            {
                if (HospitalName == "albany")
                {
                    inputColumns = new[] {
                                                "Organiztion", "HospitalName", "ItemNumber", "ItemDescription",
                                                "VendorName", "VendorCatalogNumber", "MFRName", "MFRCatalogNumber",
                                                "UOM", "Conversion", "UOMPrice", "StockStatus", "StockUOM",
                                                "IssueUOM", "UNSPSCCode", "UNSPSCCategory", "ContractNumber",
                                                "ContractType", "ContractDescription",""
                                          };

                }
                else if(HospitalName == "mercy")
                {
                    inputColumns = new[] {
                                        "","Facility","Hospital_Item_Number", "Item_Description",
                                        "Vendor Name", "Vendor Catalog Number","Manufacturer Name", "Manufacturer Catalog Number",
                                        "Purchase Unit of Measure","Conversion Factor","Price Per UOM","Status",
                                        "Stock_UOM","","","Commodity Description",
                                        "", "","Contract Number",""
                                     };
                }
                else if(HospitalName == "childrensconnecticut")
                {
                    inputColumns = new[] {
                                            "","","Item Number","Item Description",
                                            "Vendor Name","Vendor Catalog Number","Manufacturer Name","Manufacturer Catalog Number",
                                            "Purchase UOM","Purchase UOM Pack Factor","Current Price","Item Type",
                                            "Stock UOM","Issue UOM","UNSPSC Code", "UNSPSC Description",
                                            "","","",""
                                         };
                }
                else if (HospitalName == "minnesota")
                {
                    inputColumns = new[] {
                                            "","","Hospital Item Number","Item Description",
                                            "Vendor Name","Vendor Catalog Number","Manufacturer Name","Manufacturer Catalog Number",
                                            "Purchase UOM","Conversion factor","Price Per UOM","Status",
                                            "","","UNSPSC Commodity Code", "UNSPSC Commodity Description",
                                            "Contract Number","","",""
                                         };
                }
                else if (HospitalName == "nuvance")
                {
                    inputColumns = new[] {
                                            "","","Hospital Item Number","Hospital Item Description",
                                            "Vendor Name","Vendor Catalog Number","Manufacturer Name","Manufacturer Catalog Number",
                                            "Buy UOM","Buy UOM Conversion","Current Price","",
                                            "","","UNSPSC Commodity Code", "UNSPSC Commodity Description",
                                            "","","",""
                                         };
                }
                else if (HospitalName == "texas")
                {
                    inputColumns = new[] {
                                            "","","INV_ITEM_ID","DESCR254_MIXED",
                                            "NAME1","ITM_ID_VNDR","DESCR60","MFG_ITM_ID",
                                            "UNIT_OF_MEASURE","CONVERSION_RATE","PRICE_VNDR","STOCK_SYMBOL",
                                            "","","CATEGORY_CD","DESCR60_2",
                                            "MODEL","","",""
                                         };

                }
                else if(HospitalName == "uhs")
                {
                    inputColumns = new[] {
                                            "","","HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION",
                                            "VENDOR_NAME","VENDOR_CATALOG_NUMBER","MANUFACTURER_NAME","MANUFACTURER_CATALOG_NUMBER",
                                            "PACK_UOM","PURCHASE_UOM_PACK_FACTOR","CURRENT_PRICE","ITEM_TYPE",
                                            "Stock UOM","Issue UOM","UNSPSC_COMMODITY_CODE","UNSPSC_COMMODITY_DESCRIPTION",
                                            "","","",""
                                         };
                }
                else if (HospitalName == "centracare")
                {
                    inputColumns = new[] {
                                            "","","HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION",
                                            "VENDOR_NAME","VENDOR_CATALOG_NUMBER","MANUFACTURER_NAME","MANUFACTURER_CATALOG_NUMBER",
                                            "ORDER_UOM","ORDER_UOM_MULT","ORDER_PRICE","ITEM_TYPE",
                                            "","","UNSPSC_CODE","UNSPSC_DESCRIPTION",
                                            "AGREEMENT","","",""
                                         };


                }

                outputColumns = new[] {
                                        "Organization", "HospitalName", "ItemNumber", "ItemDescription",
                                        "VendorName","VendorCatalogNumber","MFRName", "MFRCatalogNumber",
                                        "UOM", "Conversion","UOMPrice","StockStatus",
                                        "StockUOM","IssueUOM","UNSPSCCode","UNSPSCCategory",
                                        "ContractNumber","ContractType","ContractDescription","OrderUOM"
                                      };
            }
            else if(mode == "po")
            {
                if(HospitalName == "albany")
                {
                    inputColumns = new[] {
                                            "Organization", "HospitalName", "PONumber", "POLineNumber",
                                            "PODate","", "DepartmentName", "ItemNumber", "ItemDescription",
                                            "VendorName", "VendorCatalogNumber", "MFRName", "MFRCatalogNumber",
                                            "UOM", "Conversion", "UOMPrice", "POQty", "POLineSpend"
                                          };
                }
                else if(HospitalName == "mercy")
                {
                    inputColumns = new[] {
                                            "","Facility","PO_Number","Line_Nbr",
                                            "PO_Date","","Department Name","Item","Description",
                                            "Vendor Name", "Ven_Item","Manufacturer Name", "Manufacturer Product Number",
                                            "PUOM","Conversion Factor","Price Per UOM","Quantity","PO Line Spend"
                                         };
                }
                else if (HospitalName == "childrensconnecticut")
                {
                    inputColumns = new[] {
                                            "","","PO Number","PO Line Number",
                                            "PO Date","","Department Name","Item Number","Item Description",
                                            "Vendor Name","Vendor Catalog Number","Manufacturer Name","Manufacturer Catalog Number",
                                            "Purchase UOM","Quantity per UOM (pack factor)","Price Paid (PO Price)","Quantity Purchased","Extended Price Paid (PO Price)"
                                         };
                }
                else if (HospitalName == "minnesota")
                {
                    inputColumns = new[] {
                                            "","","PO Number","PO Line Number",
                                            "PO Date","","Department Name","Hospital Item Number","Item Description",
                                            "Vendor Name","Vendor Product Number","Manufacturer Name","Manufacturer's Product Number",                                           "",
                                            "PO UOM","Conversion Factor","PO Price per UOM","PO Quantity","PO Line Spend"
                                         };
                }
                else if (HospitalName == "nuvance")
                {
                    inputColumns = new[] {
                                            "","Hospital Name","PO Number","PO Line Number",
                                            "PO Date","","Hospital Department Name","Hospital Item Number","Hospital Item Description",
                                            "Hospital Vendor Name","Vendor Catalog Number","Manufacturer Name","Manufacturer Catalog Number",
                                            "Purchase UOM","Quantity per UOM","Price Paid (PO Price)","Quantity Purchased","Extended Price Paid (PO Price)"
                                         };
                }
                else if(HospitalName == "texas")
                {
                    inputColumns = new[] {
                                            "","","PO_ID","LINE_NBR",
                                            "PO_DT","","DESCR","INV_ITEM_ID","DESCR254_MIXED",
                                            "NAME1","ITM_ID_VNDR","DESCR60","MFG_ITM_ID",
                                            "UNIT_OF_MEASURE","CONVERSION_RATE","PRICE_PO","QTY_PO","TCH_SPENDING_AMT"
                                         };
                }
                else if (HospitalName == "uhs")
                {
                    inputColumns = new[] {
                                            "","HOSPITAL_NAME","PO_NUMBER","PO_LINE_NUMBER",
                                            "PO_DATE","","HOSPITAL_DEPARTMENT_NAME","HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION",
                                            "VENDOR_NAME","VENDOR_CATALOG_NUMBER","MANUFACTURER_NAME","MANUFACTURER_CATALOG_NUMBER",
                                            "PURCHASE_UOM","QUANTITY_PER_UOM","PO_PRICE_PAID","QUANTITY_PURCHASED","PO_EXTENDED_PRICE_PAID"
                                         };
                }
                else if (HospitalName == "centracare")
                {
                    inputColumns = new[] {
                                            "FACILITY_NAME","","PO_NUMBER","PO_LINE_NUM",
                                            "CREATE_DATE","","DEPARTMENT_DESCRIPTION","HOSPITAL_ITEM_NUMBER","HOSPITAL_ITEM_DESCRIPTION",
                                            "VENDOR_NAME","VENDOR_PRODUCT_NUM","MANUFACTURER","MANUFACTURER_PART_NUM",
                                            "ORDER_UOM","ORDER_UOM_MULT","ORDER_UNIT_PRICE","ORDER_QUANTITY","ORDER_EXTENDED_PRICE"
                                         };
                }

                outputColumns = new[] {
                                        "Organization", "HospitalName", "PONumber", "POLineNumber",
                                        "PODate","POReference", "DepartmentName", "ItemNumber", "ItemDescription",
                                        "VendorName", "VendorCatalogNumber", "MFRName", "MFRCatalogNumber",
                                        "UOM", "Conversion", "UOMPrice", "POQty", "POLineSpend"
                                       };
            }

            Worksheet sourceSheet = workbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            Worksheet targetSheet = outputWorkbook.Worksheets[0];
            Cells targetCells = targetSheet.Cells;

            for (int i = 0; i < outputColumns.Length; i++)
            {
                targetCells[0, (byte)i].PutValue(outputColumns[i]);
            }

            int totalRows = sourceCells.MaxDataRow;

            for (int colIndex = 0; colIndex < inputColumns.Length; colIndex++)
            {
                string sourceHeaderName = inputColumns[colIndex].Trim().ToLower();

                if (!string.IsNullOrEmpty(sourceHeaderName))
                {
                    for (int col = 0; col <= sourceCells.MaxColumn; col++)
                    {
                        if (string.Equals(sourceCells[0, col].StringValue.Trim(), sourceHeaderName, StringComparison.OrdinalIgnoreCase))
                        {
                            Cell headerCell = sourceCells[0, col];
                            if (headerCell != null)
                            {
                                int sourceColNum = headerCell.Column;
                                for (int r = 1; r <= totalRows; r++)
                                {
                                    object val = sourceCells[r, sourceColNum].Value;
                                    if (val != null)
                                    {
                                        targetCells[r, colIndex].PutValue(val);
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }

            targetSheet.AutoFitColumns();

            return outputWorkbook;
        }
        #endregion
    }

    public class ConversionRequest
    {
        public string UploadedInputFileName { get; set; }
        public string InputFileName { get; set; }
        public string Delimiter { get; set; }
        public bool IsToTrimSplittedValues { get; set; } = true;
    }
}
