using NTwain.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerApp.Models
{
    public class ScannerInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public ScannerType Type { get; set; }
        public bool SupportsDuplex { get; set; }
        public string Description => $"{Name} ({Type}){(SupportsDuplex  ? " - Duplex" : "")}";
    }
}
