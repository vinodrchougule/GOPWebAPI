using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.GAT_Models
{
    public class AttributeValueConsistencyCheckModel
    {
        [Required(ErrorMessage = "Input file name is required.")]
        public string InputFileName { get; set; }
        public string UploadedInputFileName { get; set; }
        [Required(ErrorMessage = "UOM file name is required.")]
        public string UOMFileName { get; set; }
        public string MultipleValuesSeparator { get; set; } = ", ";
        public string CheckSpaceBeforeUOM { get; set; } //W - With Space, O - Without Space
        public string CheckSpaceForMultipleDimensionSeparator { get; set; } //W - With Space, O - Without Space
        public bool IsMultipleDimensionSeparatorXChecked { get; set; } = false;
        public bool IsMultipleDimensionSeparatorFWDSlashChecked { get; set; } = false;
        public bool IsMultipleDimensionSeparatorTOChecked { get; set; } = false;
        public bool IsMultipleDimensionSeparatorHyphenChecked { get; set; } = false;
        public string CheckSpaceForRangeValuesSeparator { get; set; } //W - With Space, O - Without Space
        public bool IsRangeValuesSeparatorTOChecked { get; set; } = false;
        public bool IsRangeValuesSeparatorHyphenChecked { get; set; } = false;
        public List<string> SelectedOrderedUOMList { get; set; }
        public bool IsConversionOfVAChecked { get; set; } = false;
        [Required(ErrorMessage = "UserID is required.")]
        public string UserID { get; set; }
    }
}