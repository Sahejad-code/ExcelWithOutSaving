using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Web;

namespace ExcelWithOutSaving.Models
{
    public class ExcelUploadRequestModel
    {
       
        
           public HttpPostedFileBase File { get; set; }
            public string UploadStatusLabel { get; set; }
            public DataTable DataTable { get; set; } // Add DataTable property
            public int MaxAllowedColumns { get; set; } = 5; // Set the maximum allowed columns
    }


}