using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class ConfirmDuplicateSKUsModel
    {
        public string CustomerCode { get; set; }
        public string ProjectCode { get; set; }
        public string BatchNo { get; set; }
        public string SelectedColumns { get; set; }
        public List<CIFID> CIFIDs { get; set; }
    }

    public class CIFID
    {
        public int ID { get; set; }
        public bool IsDuplicate { get; set; }
        public string DuplicateSetRemarks { get; set; }
    }
}