using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelCreate.Models
{

    public enum FileStatus
    { 
        Creating,
        Completed
    }

    public class UserFile
    {
        public int Id { get; set; }
        public string UserId { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public DateTime? CreatedAt { get; set; }
        public FileStatus FileStatus { get; set; }

        [NotMapped]
        public string GetCreatedAt
        {
            get => CreatedAt.HasValue ? CreatedAt.Value.ToShortDateString() : "-"; 
        }
    }
}
