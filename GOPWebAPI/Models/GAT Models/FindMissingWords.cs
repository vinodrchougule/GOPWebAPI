using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.GAT_Models
{
    public class FindMissingWords
    {
        public string FileName { get; set; }
        public List<string> InputColumns { get; set; }
        public List<string> OutputColumns { get; set; }
    }
}