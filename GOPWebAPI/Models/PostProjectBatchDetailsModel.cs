using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class PostProjectBatchDetailsModel
    {
        [Required]
        [StringLength(3)]
        public string CustomerCode { get; set; }

        [Required]
        [StringLength(5)]
        public string ProjectCode { get; set; }

        [Required]
        [StringLength(4)]
        public string BatchNo { get; set; }

        [Required]
        public long DuplicateCount { get; set; }

        [Required]
        public long ExceptionalCount { get; set; }

        [Required]
        public long NotProcessedCount { get; set; }

        [Required]
        public decimal QCSamplingPercentage { get; set; }

        [Required]
        public decimal QCErrorRate { get; set; }

        [Required]
        public decimal QASamplingPercentage { get; set; }

        [Required]
        public decimal QAErrorRate { get; set; }

        public string Remarks { get; set; }

        public string CAPADetails { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}