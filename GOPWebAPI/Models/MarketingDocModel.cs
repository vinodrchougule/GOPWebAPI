using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class MarketingDocModel
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string UserUploadedFileName { get; set; }
        [Required]
        public string FileName { get; set; }
        [Required]
        public string Domain { get; set; }
        [Required]
        public string DocType { get; set; }
        public string FileType { get; set; }
        public DateTime UploadedOn { get; set; }
        [Required]
        public string UserID { get; set; }
    }
}