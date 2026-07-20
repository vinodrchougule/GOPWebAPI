using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.GAT_Models
{
    public class DuplicatesCheckGenericModel
    {
        public string InputFileName { get; set; }
        public string UploadedInputFileName { get; set; }
        public string UniqueIdentifierColumnName { get; set; }
        public List<string> ColumnsToCheckForDuplicates { get; set; }
        public string DuplicatesToCheckBasedOn { get; set; }
        public double PercentageMatch { get; set; }
        public bool IsToFindDuplicatesWithExactMatch { get; set; }
        public bool IsToFindDuplicatesWithNormalizedMatch { get; set; }
        public bool IsToFindDuplicatesWithPercentageMatch { get; set; }
    }
}