using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.MRO_Dictionary
{
    public class MRODictionaryFileDataCounts
    {
        public int MRODictionaryID { get; set; }
        public DateTime UpdatedOn { get; set; }
        public string UpdatedBy { get; set; }
        public string VersionNameOrNo { get; set; }
        public string UploadedFileName { get; set; }
        public int NoOfNouns { get; set; }
        public int NoOfNounModifiers { get; set; }
        public int NoOfNounSynonyms { get; set; }
        public int NoOfNounModifierAttributes { get; set; }
        public int NoOfNounModifierAttributeEVVs { get; set; }
        public int NoOfNounModifiersMappedToUNSPSC { get; set; }
    }

    public class MRODictionaryUploadModel
    {
        public string UploadedFileName { get; set; }
        public string UploadedTempFileName { get; set; }
        public string VersionNameOrNo { get; set; }
        public string NounTableName { get; set; }
        public string NounModifierTableName { get; set; }
        public string NounSynonymTableName { get; set; }
        public string NounModifierAttributeTableName { get; set; }
        public string NounModifierAttributeValuesTableName { get; set; }
        public string NounModifierMappedUNSPSCs { get; set; }
        public string UserID { get; set; }
    }
}