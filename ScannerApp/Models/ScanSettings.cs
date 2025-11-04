using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScannerApp.Models
{
    public class ScanSettings
    {
        public string SavePath { get; set; } = @"C:\ScannedImages\";
        public bool UseDuplex { get; set; }
        public ColorMode ColorMode { get; set; } = ColorMode.Color;
        public int Resolution { get; set; } = 300;
        public ImageFormat Format { get; set; } = ImageFormat.Png;
        public bool ShowUI { get; set; } = false;
    }
}
