using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ProductionDownload
    {
        public long ProductionAllocationID { get; set; }

        public string Activities { get; set; }

        public string ProductionUser { get; set; }
    }

    public class ProductionUpload
    {
        [Required]
        public long ProductionAllocationID { get; set; }

        [Required]
        [StringLength(4000)]
        public string Activities { get; set; }

        [Required]
        [StringLength(100)]
        public string UploadedFileName { get; set; }

        [Required]
        [Range(0, 23)]
        public int WorkedHours { get; set; }

        [Required]
        [Range(0, 59)]
        public int WorkedMinutes { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class ProductionExistingUpload
    {
        public long ProductionUploadID { get; set; }

        public DateTime UploadedOn { get; set; }

        public string UploadedByUserName { get; set; }

        public string Activities { get; set; }

        public int NoOfSKUs { get; set; }

        public string UploadedFileName { get; set; }

        public int IsProductionCompletedCountDownloaded { get; set; }
    }

    public class ProductionRowData
    {
        [Required]
        public long ProductionAllocationID { get; set; }

        [Required]
        public string UniqueColumnValue { get; set; }

        public List<ProductionColumnNameAndValue> ProductionColumnValues { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class ProductionColumnNameAndValue
    {
        public string ColumnName { get; set; }
        public string ColumnValue { get; set; }
    }

    public class ProductionItem
    {
        public long ProductionAllocationID { get; set; }

        public long ProductionItemID { get; set; }
        public long QCItemID { get; set; }
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public long? PreviousProductionItemID { get; set; }

        public long? NextProductionItemID { get; set; }

        [Required]
        [StringLength(100)]
        public string UniqueID { get; set; }

        public string ShortDescription { get; set; }
        public string LongDescription { get; set; }
        public string UOM { get; set; }
        public string NewShortDescription { get; set; }
        public string NewLongDescription { get; set; }
        public string MissingWords { get; set; }
        public string MFRName { get; set; }
        public string MFRPN { get; set; }
        public string VendorName { get; set; }
        public string VendorPN { get; set; }
        public string Noun { get; set; }
        public string Modifier { get; set; }
        public string Status { get; set; }
        public string Level { get; set; }
        public string MFRName1 { get; set; }
        public string MFRPN1 { get; set; }
        public string MFRName2 { get; set; }
        public string MFRPN2 { get; set; }
        public string MFRName3 { get; set; }
        public string MFRPN3 { get; set; }
        public string VendorName1 { get; set; }
        public string VendorPN1 { get; set; }
        public string VendorName2 { get; set; }
        public string VendorPN2 { get; set; }
        public string VendorName3 { get; set; }
        public string VendorPN3 { get; set; }
        public string AdditionalInfo { get; set; }
        public string AdditionalInfoFromWeb { get; set; }
        public string UNSPSCCode { get; set; }
        public string UNSPSCCategory { get; set; }
        public string WebRefURL1 { get; set; }
        public string WebRefURL2 { get; set; }
        public string WebRefURL3 { get; set; }
        public string PDFURL { get; set; }
        public string Remarks { get; set; }
        public string Query { get; set; }
        public string IsMovedToQC { get; set; }
        public string UNSPSCVersion { get; set; }
        public string ProductionUser { get; set; }
        public string QCUser { get; set; }
        public string QCStatus { get; set; }
        public string QCTestNo { get; set; }
        public int TotalRowsCount { get; set; }
        public string CustomColumnName1 { get; set; }
        public string CustomColumnName1Value { get; set; }
        public string CustomColumnName2 { get; set; }
        public string CustomColumnName2Value { get; set; }
        public string CustomColumnName3 { get; set; }
        public string CustomColumnName3Value { get; set; }
        public string Application { get; set; }
        public string DWG { get; set; }
        public string ItemNo { get; set; }
        public string OtherNo { get; set; }
        public string POS { get; set; }
        public string SerialNo { get; set; }
        public string KKSCode { get; set; }
        public string AssemblyOrPart { get; set; }      //A or P
        public string BOM { get; set; }
        public string GreenItems { get; set; }          //Y or N
        public int AllocatedCount { get; set; }
        public int PendingCount { get; set; }
        public int ProcessedCount { get; set; }
        public int QCApprovedCount { get; set; }
        public int RejectedCount { get; set; }
        public int QueryCount { get; set; }
        public DateTime UpdatedOn { get; set; }
        public string AttributeName1 { get; set; }
        public string AttributeValue1 { get; set; }
        public string AttributeName2 { get; set; }
        public string AttributeValue2 { get; set; }
        public string AttributeName3 { get; set; }
        public string AttributeValue3 { get; set; }
        public string AttributeName4 { get; set; }
        public string AttributeValue4 { get; set; }
        public string AttributeName5 { get; set; }
        public string AttributeValue5 { get; set; }
        public string AttributeName6 { get; set; }
        public string AttributeValue6 { get; set; }
        public string AttributeName7 { get; set; }
        public string AttributeValue7 { get; set; }
        public string AttributeName8 { get; set; }
        public string AttributeValue8 { get; set; }
        public string AttributeName9 { get; set; }
        public string AttributeValue9 { get; set; }
        public string AttributeName10 { get; set; }
        public string AttributeValue10 { get; set; }
        public string AttributeName11 { get; set; }
        public string AttributeValue11 { get; set; }
        public string AttributeName12 { get; set; }
        public string AttributeValue12 { get; set; }
        public string AttributeName13 { get; set; }
        public string AttributeValue13 { get; set; }
        public string AttributeName14 { get; set; }
        public string AttributeValue14 { get; set; }
        public string AttributeName15 { get; set; }
        public string AttributeValue15 { get; set; }
        public string AttributeName16 { get; set; }
        public string AttributeValue16 { get; set; }
        public string AttributeName17 { get; set; }
        public string AttributeValue17 { get; set; }
        public string AttributeName18 { get; set; }
        public string AttributeValue18 { get; set; }
        public string AttributeName19 { get; set; }
        public string AttributeValue19 { get; set; }
        public string AttributeName20 { get; set; }
        public string AttributeValue20 { get; set; }
        public string AttributeName21 { get; set; }
        public string AttributeValue21 { get; set; }
        public string AttributeName22 { get; set; }
        public string AttributeValue22 { get; set; }
        public string AttributeName23 { get; set; }
        public string AttributeValue23 { get; set; }
        public string AttributeName24 { get; set; }
        public string AttributeValue24 { get; set; }
        public string AttributeName25 { get; set; }
        public string AttributeValue25 { get; set; }
        public string AttributeName26 { get; set; }
        public string AttributeValue26 { get; set; }
        public string AttributeName27 { get; set; }
        public string AttributeValue27 { get; set; }
        public string AttributeName28 { get; set; }
        public string AttributeValue28 { get; set; }
        public string AttributeName29 { get; set; }
        public string AttributeValue29 { get; set; }
        public string AttributeName30 { get; set; }
        public string AttributeValue30 { get; set; }
        public string AttributeName31 { get; set; }
        public string AttributeValue31 { get; set; }
        public string AttributeName32 { get; set; }
        public string AttributeValue32 { get; set; }
        public string AttributeName33 { get; set; }
        public string AttributeValue33 { get; set; }
        public string AttributeName34 { get; set; }
        public string AttributeValue34 { get; set; }
        public string AttributeName35 { get; set; }
        public string AttributeValue35 { get; set; }
        public string AttributeName36 { get; set; }
        public string AttributeValue36 { get; set; }
        public string AttributeName37 { get; set; }
        public string AttributeValue37 { get; set; }
        public string AttributeName38 { get; set; }
        public string AttributeValue38 { get; set; }
        public string AttributeName39 { get; set; }
        public string AttributeValue39 { get; set; }
        public string AttributeName40 { get; set; }
        public string AttributeValue40 { get; set; }
        public string AttributeName41 { get; set; }
        public string AttributeValue41 { get; set; }
        public string AttributeName42 { get; set; }
        public string AttributeValue42 { get; set; }
        public string AttributeName43 { get; set; }
        public string AttributeValue43 { get; set; }
        public string AttributeName44 { get; set; }
        public string AttributeValue44 { get; set; }
        public string AttributeName45 { get; set; }
        public string AttributeValue45 { get; set; }
        public string AttributeName46 { get; set; }
        public string AttributeValue46 { get; set; }
        public string AttributeName47 { get; set; }
        public string AttributeValue47 { get; set; }
        public string AttributeName48 { get; set; }
        public string AttributeValue48 { get; set; }
        public string AttributeName49 { get; set; }
        public string AttributeValue49 { get; set; }
        public string AttributeName50 { get; set; }
        public string AttributeValue50 { get; set; }
        public List<ItemAttribute> ItemAttributes { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class ItemAttribute
    {
        public string AttributeName { get; set; }
        public string AttributeValue { get; set; }
        public string QCAttributeValue { get; set; }
        public string QCAttributeValueComments { get; set; }

    }

    public class NounModifier
    {
        public string Noun { get; set; }
        public string Modifier { get; set; }
    }

    public class NounModifierAttribute
    {
        public string AttributeName { get; set; }
    }

    public class ProjectNounModifierAttribute
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string Noun { get; set; }
        public string Modifier { get; set; }
        public string AttributeName { get; set;} 
    }

    public class MFRVendorPNProjectModel
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string UniqueID { get; set; }
        public string MFRPN1 { get; set; }
        public string MFRPN2 { get; set; }
        public string MFRPN3 { get; set; }
        public string VendorPN1 { get; set; }
        public string VendorPN2 { get; set; }
        public string VendorPN3 { get; set; }
    }

    public class DuplicateToFindOnColumnsModel
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public List<DuplicateToFindOnColumns> ColumnNames { get; set; }
    }
    
    public class DuplicateToFindOnColumns
    {
        public string ColumnName { get; set; }
    }

    public class MoveToQCModel
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public long ProductionAllocationID { get; set; }
        public List<ProductionItemIDModel> ProductionItemIDs { get; set; }
        public string UserID { get; set; }
    }

    public class ProductionItemIDModel
    {
        public long ProductionItemID { get; set; }
    }

    public class ProjectOverallStatusCount
    {
        public int Volume { get; set; }
        public int AllocatedCount { get; set; }
        public int YetToAllocate { get; set; }
        public int Processed { get; set; }
        public int QCApproved { get; set; }
        public int Query { get; set; }
        public decimal ProductionCompletedPercentage { get; set; }
        public decimal QCCompletedPercentage { get; set; }
    }

    public class ProjectStatusUserDateCount
    {
        public string User {  get; set; }
        public DateTime UpdatedOn { get; set; }
        public int Count { get; set; }
    }

    public class ProjectStatusSKUs
    {
        public string UniqueID { get; set; }
        public string Level { get; set; }
        public string User { get; set; }
        public DateTime UpdatedOn { get; set; }
    }

    public class TripleResultsDto
    {
        public ProjectOverallStatusCount projectOverallStatusCount { get; set; }
        public List<ProjectStatusUserDateCount> projectStatusUserDateCountList { get; set; }
        public List<ProjectStatusSKUs> projectStatusSKUsList { get; set; }
    }

    public class MRORefDBPartDetails
    {
        public long MRORefDBID { get; set; }
        public string MFRName1 { get; set; }
        public string MFRPN1 { get; set; }
        public string MFRName2 { get; set; }
        public string MFRPN2 { get; set; }
        public string RefURL1 { get; set; }
        public string RefURL2 { get; set; }
        public string Noun { get; set; }
        public string Modifier { get; set; }
        public string AttributeName1 { get; set; }
        public string AttributeValue1 { get; set; }
        public string AttributeName2 { get; set; }
        public string AttributeValue2 { get; set; }
        public string AttributeName3 { get; set; }
        public string AttributeValue3 { get; set; }
        public string AttributeName4 { get; set; }
        public string AttributeValue4 { get; set; }
        public string AttributeName5 { get; set; }
        public string AttributeValue5 { get; set; }
        public string AttributeName6 { get; set; }
        public string AttributeValue6 { get; set; }
        public string AttributeName7 { get; set; }
        public string AttributeValue7 { get; set; }
        public string AttributeName8 { get; set; }
        public string AttributeValue8 { get; set; }
        public string AttributeName9 { get; set; }
        public string AttributeValue9 { get; set; }
        public string AttributeName10 { get; set; }
        public string AttributeValue10 { get; set; }
        public string AttributeName11 { get; set; }
        public string AttributeValue11 { get; set; }
        public string AttributeName12 { get; set; }
        public string AttributeValue12 { get; set; }
        public string AttributeName13 { get; set; }
        public string AttributeValue13 { get; set; }
        public string AttributeName14 { get; set; }
        public string AttributeValue14 { get; set; }
        public string AttributeName15 { get; set; }
        public string AttributeValue15 { get; set; }
        public string AttributeName16 { get; set; }
        public string AttributeValue16 { get; set; }
        public string AttributeName17 { get; set; }
        public string AttributeValue17 { get; set; }
        public string AttributeName18 { get; set; }
        public string AttributeValue18 { get; set; }
        public string AttributeName19 { get; set; }
        public string AttributeValue19 { get; set; }
        public string AttributeName20 { get; set; }
        public string AttributeValue20 { get; set; }
        public List<ItemAttribute> ItemAttributes { get; set; }
    }

}