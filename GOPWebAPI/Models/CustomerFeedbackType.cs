using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models
{
    public class CustomerFeedbackType
    {
        public int SlNo { get; set; }

        [Key]
        public int CustomerFeedbackTypeID { get; set; }

        [Required]
        [StringLength(50)]
        public string FeedbackType { get; set; }

        [Required]
        public bool IsActive { get; set; }

        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }
}