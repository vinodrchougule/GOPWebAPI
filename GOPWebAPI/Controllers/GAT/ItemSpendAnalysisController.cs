using Aspose.Cells;
using GOPWebAPI.BLL;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models;
using GOPWebAPI.Models.GAT_Models;
using Swashbuckle.Swagger;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.DynamicData;
using System.Web.Http;

namespace GOPWebAPI.Controllers.GAT
{
    [RoutePrefix("api/ItemSpendAnalysis")]
    public class ItemSpendAnalysisController : ApiController
    {
        #region Validate the input uploaded file
        [HttpPost]
        [Route("ValidateInputFile")]
        public HttpResponseMessage ValidateInputFile(string FileName)
        {
            string UploadedFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + FileName);

            try
            {
                if (HttpContext.Current == null)
                    throw new HttpResponseException(HttpStatusCode.Unauthorized);

                string ext = Path.GetExtension(FileName).ToLower();

                if (ext != ".xlsx")
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Please upload a valid input file of xlsx format only");

                Workbook wbIF = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wbIF.Open(UploadedFilepath);
                Worksheet ws = wbIF.Worksheets[0];

                if (ws.Cells[0, 0].StringValue.Trim().ToUpper() != "SCA CODE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "First Column with heading 'SCA Code' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 1].StringValue.Trim().ToUpper() != "GPO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Second Column with heading 'GPO' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 2].StringValue.Trim().ToUpper() != "HEALTH SYSTEM NAME")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Third Column with heading 'Health System Name' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 3].StringValue.Trim().ToUpper() != "FACILITY NAME")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourth Column with heading 'Facility Name' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 4].StringValue.Trim().ToUpper() != "POC")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fifth Column with heading 'POC' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 5].StringValue.Trim().ToUpper() != "DEPT / AREA")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Sixth Column with heading 'Dept / Area' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 6].StringValue.Trim().ToUpper() != "COMMODITY TITLE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Seventh Column with heading 'Commodity Title' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 7].StringValue.Trim().ToUpper() != "ITEM NO")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Eighth Column with heading 'Item No' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 8].StringValue.Trim().ToUpper() != "DESCRIPTION")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Ninth Column with heading 'Description' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 9].StringValue.Trim().ToUpper() != "VENDOR")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Tenth Column with heading 'Vendor' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 10].StringValue.Trim().ToUpper() != "VENDOR CAT")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Eleventh Column with heading 'Vendor Cat' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 11].StringValue.Trim().ToUpper() != "MFR NAME")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Twelth Column with heading 'MFR Name' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 12].StringValue.Trim().ToUpper() != "MFR CAT")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Thirteenth Column with heading 'MFR Cat' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 13].StringValue.Trim().ToUpper() != "UOM")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fourteenth Column with heading 'UOM' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 14].StringValue.Trim().ToUpper() != "PKG")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Fifteenth Column with heading 'PKG' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 15].StringValue.Trim().ToUpper() != "HOSPITAL QTY")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Sixteenth Column with heading 'Hospital Qty' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 16].StringValue.Trim().ToUpper() != "HOSPITAL PRICE")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Seventeenth Column with heading 'Hospital Price' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells[0, 17].StringValue.Trim().ToUpper() != "HOSPITAL VOLUME")
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Eighteenth Column with heading 'Hospital Volume' not found in first row of input file first worksheet. Please select a valid file.");
                }

                if (ws.Cells.Rows.Count <= 1)
                {
                    File.Delete(UploadedFilepath);
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "Input File first Worksheet has no data rows.");
                }

                return Request.CreateResponse(HttpStatusCode.OK, "Input file validated successfully.");
            }
            catch (Exception ex)
            {
                if (File.Exists(UploadedFilepath))
                    File.Delete(UploadedFilepath);

                ExceptionLogging.SendExceptionToDB(ex);
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.Message);
            }
        }
        #endregion

        #region Do Item Spend Analysis and Write the Output to file
        [HttpPost]
        [Route("DoItemSpendAnalysis")]
        public HttpResponseMessage DoItemSpendAnalysis(string InputFileName, string UploadedInputFileName)
        {
            try
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

                string UploadedInputFilepath = HttpContext.Current.Server.MapPath(@"\temp\" + InputFileName);

                Workbook wi = new Workbook();
                Aspose.Cells.License l = new Aspose.Cells.License();
                l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wi.LoadData(UploadedInputFilepath);

                Worksheet swsi = wi.Worksheets[0];
                int MaxDataRows = swsi.Cells.MaxRow;

                #region Reading Input Worksheet and Writing to List
                List<MBInput> MBInputList = new List<MBInput>();
                for (int i = 1; i <= MaxDataRows; i++)
                {
                    MBInput mi = new MBInput();
                    mi.SCACode = swsi.Cells[i, 0].StringValue.Trim();
                    mi.GPO = swsi.Cells[i, 1].StringValue.Trim();
                    mi.HealthSystemName = swsi.Cells[i, 2].StringValue.Trim();
                    mi.FacilityName = swsi.Cells[i, 3].StringValue.Trim();
                    mi.POC = swsi.Cells[i, 4].StringValue.Trim();
                    mi.DeptArea = swsi.Cells[i, 5].StringValue.Trim();
                    mi.CommodityTitle = swsi.Cells[i, 6].StringValue.Trim();
                    mi.ItemNo = swsi.Cells[i, 7].StringValue.Trim();
                    mi.Description = swsi.Cells[i, 8].StringValue.Trim();
                    mi.Vendor = swsi.Cells[i, 9].StringValue.Trim();
                    mi.VendorCat = swsi.Cells[i, 10].StringValue.Trim();
                    mi.MFRName = swsi.Cells[i, 11].StringValue.Trim();
                    mi.MFRCat = swsi.Cells[i, 12].StringValue.Trim();
                    mi.UOM = swsi.Cells[i, 13].StringValue.Trim();
                    mi.PKG = swsi.Cells[i, 14].StringValue.Trim();
                    mi.HospitalQty = swsi.Cells[i, 15].StringValue.Trim();
                    mi.HospitalPrice = swsi.Cells[i, 16].StringValue.Trim();
                    mi.HospitalVolume = swsi.Cells[i, 17].StringValue.Trim();
                    MBInputList.Add(mi);
                }

                MBInputList = MBInputList.OrderBy(d => d.DeptArea).ThenBy(c => c.CommodityTitle).ThenByDescending(v => v.decimalHospitalVolume).ToList();
                #endregion

                string tempPath = System.IO.Path.GetTempPath();
                string filename = tempPath + "Output_" + UploadedInputFileName;
                wi.Save(filename);

                Workbook wo = new Workbook();
                Aspose.Cells.License l2 = new Aspose.Cells.License();
                l2.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
                wo.LoadData(filename);
                Worksheet swso = wo.Worksheets[0];

                #region Writing Headers to Output Worksheet
                for (int c = 0; c <= 17; c++)
                {
                    swso.Cells[0, c].PutValue(swsi.Cells[0, c].StringValue.Trim());
                    swso.Cells[0, c].SetStyle(swsi.Cells[0, c].GetStyle());
                }
                #endregion

                #region Writing Data To Output Worksheet
                int or = 1;
                foreach (MBInput mi in MBInputList)
                {
                    swso.Cells[or, 0].PutValue(mi.SCACode);
                    swso.Cells[or, 1].PutValue(mi.GPO);
                    swso.Cells[or, 2].PutValue(mi.HealthSystemName);
                    swso.Cells[or, 3].PutValue(mi.FacilityName);
                    swso.Cells[or, 4].PutValue(mi.POC);
                    swso.Cells[or, 5].PutValue(mi.DeptArea);
                    swso.Cells[or, 6].PutValue(mi.CommodityTitle);
                    swso.Cells[or, 7].PutValue(mi.ItemNo);
                    swso.Cells[or, 8].PutValue(mi.Description);
                    swso.Cells[or, 9].PutValue(mi.Vendor);
                    swso.Cells[or, 10].PutValue(mi.VendorCat);
                    swso.Cells[or, 11].PutValue(mi.MFRName);
                    swso.Cells[or, 12].PutValue(mi.MFRCat);
                    swso.Cells[or, 13].PutValue(mi.UOM);
                    swso.Cells[or, 14].PutValue(mi.PKG);
                    swso.Cells[or, 15].PutValue(mi.HospitalQty);
                    swso.Cells[or, 16].PutValue(mi.HospitalPrice);
                    swso.Cells[or, 17].PutValue(mi.HospitalVolume);
                    or++;
                }
                #endregion

                #region Calculating 90-10 starts
                Aspose.Cells.Style styleNavyBlueFillColor = swso.Cells[0, 17].GetStyle();
                swso.Cells[0, 18].PutValue("90-10");
                swso.Cells[0, 18].SetStyle(styleNavyBlueFillColor);

                #region Creating a list of distinct departments
                string DepartmentName = string.Empty;
                List<string> Departments = new List<string>();
                for (int r = 1; r <= MaxDataRows; r++)
                {
                    DepartmentName = swso.Cells[r, 5].StringValue.Trim();
                    if (!Departments.Contains(DepartmentName))
                        Departments.Add(DepartmentName);
                }
                #endregion

                #region Calculating Department / Area Total of each Department / Area
                DataTable dtDepartment = new DataTable();
                dtDepartment.Columns.Add("Department");
                dtDepartment.Columns.Add("Total");

                string HospitalVolume = string.Empty;
                decimal DepartmentTotal = 0;
                foreach (string department in Departments)
                {
                    DepartmentTotal = 0;
                    for (int ir = 1; ir <= MaxDataRows; ir++)
                    {
                        if (swso.Cells[ir, 5].StringValue.Trim().ToUpper() == department.ToUpper())
                        {
                            HospitalVolume = swso.Cells[ir, 17].StringValue.Replace("$", "").Trim();
                            HospitalVolume = HospitalVolume.Replace("-", "0");
                            DepartmentTotal = DepartmentTotal + Convert.ToDecimal(HospitalVolume);
                        }
                    }

                    DataRow dr = dtDepartment.NewRow();
                    dr["Department"] = department;
                    dr["Total"] = DepartmentTotal;
                    dtDepartment.Rows.Add(dr);
                }
                #endregion

                #region Creating list of distinct commodities of each department
                List<clsCommodity> Commodities = new List<clsCommodity>();
                string CommodityName = string.Empty, Department = string.Empty;
                for (int ir = 1; ir <= MaxDataRows; ir++)
                {
                    Department = swso.Cells[ir, 5].StringValue.Trim();
                    CommodityName = swso.Cells[ir, 6].StringValue.Trim();
                    if (Commodities.Count(c => c.Department.Trim().ToUpper() == Department.Trim().ToUpper() && c.Commodity.Trim().ToUpper() == CommodityName.Trim().ToUpper()) == 0)
                    {
                        clsCommodity commodity = new clsCommodity();
                        commodity.Department = Department;
                        commodity.Commodity = CommodityName;
                        Commodities.Add(commodity);
                    }
                }
                #endregion

                #region Calculating Hospital Volume Total for each Commodity of each department
                decimal CommodityTotal = 0;
                List<clsCommodity> CommodityList = new List<clsCommodity>();
                foreach (clsCommodity commodity in Commodities)
                {
                    CommodityTotal = 0;
                    for (int ir = 1; ir <= MaxDataRows; ir++)
                    {
                        Department = swso.Cells[ir, 5].StringValue.Trim();
                        CommodityName = swso.Cells[ir, 6].StringValue.Trim();
                        HospitalVolume = swso.Cells[ir, 17].StringValue.Replace("$", "").Trim();
                        HospitalVolume = HospitalVolume.Replace("-", "0");
                        if (commodity.Department.ToUpper() == Department.ToUpper() && commodity.Commodity.ToUpper() == CommodityName.ToUpper())
                            CommodityTotal = CommodityTotal + Convert.ToDecimal(HospitalVolume);
                    }

                    clsCommodity Commodity = new clsCommodity();
                    Commodity.Department = commodity.Department;
                    Commodity.Commodity = commodity.Commodity;
                    Commodity.Total = CommodityTotal;

                    CommodityList.Add(Commodity);
                }
                #endregion

                #region Sort the List on descending order of Commodity Total
                CommodityList = CommodityList.OrderByDescending(o => o.Total).ToList();
                #endregion

                #region Adding Top 90% Commodities to list
                List<clsCommodity> Top90PercentCommodityList = new List<clsCommodity>();
                decimal sumOfCommodityTotal = 0, DepartmentTotal90Percent = 0, NinetyPercentValue = 0.9M;
                for (int i = 0; i < dtDepartment.Rows.Count; i++)
                {
                    DepartmentName = dtDepartment.Rows[i]["Department"].ToString();
                    DepartmentTotal = Convert.ToDecimal(dtDepartment.Rows[i]["Total"]);

                    sumOfCommodityTotal = 0;
                    DepartmentTotal90Percent = DepartmentTotal * NinetyPercentValue;
                    foreach (clsCommodity commodity in CommodityList)
                    {
                        if (commodity.Department == DepartmentName)
                        {
                            sumOfCommodityTotal = sumOfCommodityTotal + commodity.Total;
                            if (sumOfCommodityTotal < DepartmentTotal90Percent)
                            {
                                clsCommodity TopCommodity = new clsCommodity();
                                TopCommodity.Department = commodity.Department;
                                TopCommodity.Commodity = commodity.Commodity;
                                TopCommodity.Total = commodity.Total;
                                TopCommodity.sumOfCommodityTotal = sumOfCommodityTotal;
                                Top90PercentCommodityList.Add(TopCommodity);
                            }
                            else
                            {
                                clsCommodity TopCommodity = new clsCommodity();
                                TopCommodity.Department = commodity.Department;
                                TopCommodity.Commodity = commodity.Commodity;
                                TopCommodity.Total = commodity.Total;
                                TopCommodity.sumOfCommodityTotal = sumOfCommodityTotal;
                                Top90PercentCommodityList.Add(TopCommodity);
                                break;
                            }
                        }
                    }
                }
                #endregion

                #region Write 90-10 column data
                for (int row = 1; row <= MaxDataRows; row++)
                {
                    DepartmentName = swso.Cells[row, 5].StringValue.Trim();
                    CommodityName = swso.Cells[row, 6].StringValue.Trim();
                    if (Top90PercentCommodityList.Count(ncl => ncl.Department.ToUpper() == DepartmentName.ToUpper() && ncl.Commodity.ToUpper() == CommodityName.ToUpper()) > 0)
                        swso.Cells[row, 18].PutValue("Yes");
                    else
                        swso.Cells[row, 18].PutValue("No");
                }
                #endregion
                #endregion

                #region Calculating 80-20 starts
                swso.Cells[0, 19].PutValue("80-20");
                swso.Cells[0, 19].SetStyle(styleNavyBlueFillColor);

                #region Creating a list of "90-10 - Yes" for each department and commodity
                List<clsCommodity> NinetyTenYesDepartmentCommodities = new List<clsCommodity>();
                string NinetyTen = string.Empty;
                for (int row = 1; row <= MaxDataRows; row++)
                {
                    DepartmentName = swso.Cells[row, 5].StringValue.Trim();
                    CommodityName = swso.Cells[row, 6].StringValue.Trim();
                    HospitalVolume = swso.Cells[row, 17].StringValue.Replace("$", "").Trim();
                    HospitalVolume = HospitalVolume.Replace("-", "0");
                    NinetyTen = swso.Cells[row, 18].StringValue.Trim();
                    if (NinetyTen.ToUpper() == "YES")
                    {
                        clsCommodity NinetyTenCommodity = new clsCommodity();
                        NinetyTenCommodity.Department = DepartmentName;
                        NinetyTenCommodity.Commodity = CommodityName;
                        NinetyTenCommodity.HospitalVolume = Convert.ToDecimal(HospitalVolume);
                        NinetyTenYesDepartmentCommodities.Add(NinetyTenCommodity);
                    }
                }
                #endregion

                #region Creating a list of "90-10 - Yes" commodities Total with EightyPercent
                var NinetyTenYesDepartmentCommoditiesWithTotal = NinetyTenYesDepartmentCommodities
                                                        .GroupBy(g => new { g.Department, g.Commodity })
                                                        .Select(s => new { Department = s.Key.Department, Commodity = s.Key.Commodity, Total = s.Sum(c => c.HospitalVolume) });

                List<clsCommodity> NinetyTenYesDepartmentCommodityTotalList = new List<clsCommodity>();
                decimal EightyPercent = 0.8M;
                foreach (var c in NinetyTenYesDepartmentCommoditiesWithTotal)
                {
                    clsCommodity commodity = new clsCommodity();
                    commodity.Department = c.Department.ToString();
                    commodity.Commodity = c.Commodity;
                    commodity.Total = c.Total;
                    commodity.EightyPercentOfCommodityTotal = c.Total * EightyPercent;
                    NinetyTenYesDepartmentCommodityTotalList.Add(commodity);
                }
                NinetyTenYesDepartmentCommodityTotalList = NinetyTenYesDepartmentCommodityTotalList.OrderBy(d => d.Department).ThenBy(c => c.Commodity).ThenByDescending(o => o.Total).ToList();
                #endregion

                #region Writing Yes/No to Department Commodity lines
                decimal EightyPercentOfCommodityTotal = 0;
                bool IsCommodityTotalExceedingFirstTime = true;
                foreach (clsCommodity commodity in NinetyTenYesDepartmentCommodityTotalList)
                {
                    EightyPercentOfCommodityTotal = commodity.EightyPercentOfCommodityTotal;
                    CommodityTotal = 0;
                    IsCommodityTotalExceedingFirstTime = true;
                    for (int row = 1; row <= MaxDataRows; row++)
                    {
                        DepartmentName = swso.Cells[row, 5].StringValue.Trim();
                        CommodityName = swso.Cells[row, 6].StringValue.Trim();
                        if (swso.Cells[row, 18].StringValue.Trim().ToUpper() == "YES")
                        {
                            if (commodity.Department.ToUpper() == DepartmentName.ToUpper() && commodity.Commodity.ToUpper() == CommodityName.ToUpper())
                            {
                                HospitalVolume = swso.Cells[row, 17].StringValue.Replace("$", "").Trim();
                                HospitalVolume = HospitalVolume.Replace("-", "0");
                                CommodityTotal = CommodityTotal + Convert.ToDecimal(HospitalVolume);
                                if (CommodityTotal <= EightyPercentOfCommodityTotal)
                                {
                                    swso.Cells[row, 19].PutValue("Yes");
                                    IsCommodityTotalExceedingFirstTime = true;
                                }
                                else
                                {
                                    if (IsCommodityTotalExceedingFirstTime)
                                    {
                                        swso.Cells[row, 19].PutValue("Yes");
                                        IsCommodityTotalExceedingFirstTime = false;
                                    }
                                    else
                                    {
                                        swso.Cells[row, 19].PutValue("No");
                                        IsCommodityTotalExceedingFirstTime = false;
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
                #endregion

                #region Creating Top80 Worksheet and Writing Top80 Percent Data - 80-20 - Yes
                bool IsTop80WorksheetExists = false;
                foreach (Worksheet ws in wo.Worksheets)
                {
                    if (ws.Name.Trim().ToUpper() == "TOP 80")
                    {
                        IsTop80WorksheetExists = true;
                        break;
                    }
                }

                if (!IsTop80WorksheetExists)
                    wo.Worksheets.Add("Top 80");

                Worksheet wsTop80 = wo.Worksheets["Top 80"];
                #region Writing Headers from swso worksheet
                int MaxDataColumns = swso.Cells.MaxColumn;
                for (int ic = 0; ic <= MaxDataColumns - 2; ic++)        //Excluding 90-10 & 80-20 columns
                {
                    wsTop80.Cells[0, ic].PutValue(swso.Cells[0, ic].StringValue);
                    wsTop80.Cells[0, ic].SetStyle(swso.Cells[0, ic].GetStyle());
                }
                #endregion

                #region Writing Data Rows 'Yes' from 80-20 column of swso
                int ter = 0;
                for (int ir = 1; ir <= MaxDataRows; ir++)
                {
                    if (swso.Cells[ir, 19].StringValue.Trim().ToUpper() == "YES")
                    {
                        ter++;
                        for (int ic = 0; ic <= MaxDataColumns - 2; ic++)
                        {
                            wsTop80.Cells[ter, ic].PutValue(swso.Cells[ir, ic].StringValue);
                            wsTop80.Cells[ter, ic].SetStyle(swso.Cells[ir, ic].GetStyle());
                        }
                    }
                }
                #endregion
                wsTop80.AutoFitColumns();
                #endregion

                #region Creating Bottom20 Worksheet and Writing Bottom20 or Blank Data - 80-20 - No/Blank
                bool IsBottom20WorksheetExists = false;
                foreach (Worksheet ws in wo.Worksheets)
                {
                    if (ws.Name.Trim().ToUpper() == "BOTTOM 20")
                    {
                        IsBottom20WorksheetExists = true;
                        break;
                    }
                }

                if (!IsBottom20WorksheetExists)
                    wo.Worksheets.Add("Bottom 20");

                Worksheet wsBottom20 = wo.Worksheets["Bottom 20"];
                #region Writing Headers from swso worksheet
                for (int ic = 0; ic <= MaxDataColumns - 2; ic++)        //Excluding 90-10 & 80-20 columns
                {
                    wsBottom20.Cells[0, ic].PutValue(swso.Cells[0, ic].StringValue);
                    wsBottom20.Cells[0, ic].SetStyle(swso.Cells[0, ic].GetStyle());
                }
                #endregion

                #region Writing Data Rows 'No/Blank' from 80-20 column of swso
                int btr = 0;
                for (int ir = 1; ir <= MaxDataRows; ir++)
                {
                    if (swso.Cells[ir, 19].StringValue.Trim().ToUpper() == "NO" || string.IsNullOrEmpty(swso.Cells[ir, 19].StringValue.Trim()))
                    {
                        btr++;
                        for (int ic = 0; ic <= MaxDataColumns - 2; ic++)
                        {
                            wsBottom20.Cells[btr, ic].PutValue(swso.Cells[ir, ic].StringValue);
                            wsBottom20.Cells[btr, ic].SetStyle(swso.Cells[ir, ic].GetStyle());
                        }
                    }
                }
                #endregion
                wsBottom20.AutoFitColumns();
                #endregion

                string filename1 = tempPath + "Output_" + UploadedInputFileName;
                wo.Save(filename1);

                if (File.Exists(UploadedInputFilepath))
                    File.Delete(UploadedInputFilepath);

                return Request.CreateResponse(HttpStatusCode.OK, filename);
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
    }
}
