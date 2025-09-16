using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerApp.Models
{
    public class ScanResult
    {
        public bool Success { get; set; }
        public string? FilePath { get; set; }
        public string? ErrorMessage { get; set; }
        public DateTime ScannedAt { get; set; } = DateTime.Now;
    }
}
