using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ScannerApp.Models
{

    public class OcrResponse
    {
        public OcrData Data { get; set; }
        public string Status { get; set; }
        public string FrontPath { get; set; }
        public string BackPath { get; set; }
        public string PassportPath { get; set; }
        public string SessionId { get; set; }
        public int TempPersonId { get; set; }
        public bool ActiveStatus { get; set; }//Refer to active status in visitor table
        public string ActiveStatusNote { get; set; }//Refer to active status note in visitor table
    }
    public class OcrData
    {
        public string Address { get; set; }
        public string BirthDate { get; set; }
        public string Demo { get; set; }
        public string ExpiryDate { get; set; }
        public string FirstName { get; set; }
        public string FullName { get; set; }
        public string Gender { get; set; }
        public string Governorate { get; set; }
        public string ID { get; set; }
        public string Job { get; set; }
        public string NationalID { get; set; }
        public string PassportNumber { get; set; }
        public string IssuingCountry { get; set; }
        public string LastName { get; set; }
        public string Serial { get; set; }
        public string Nationality { get; set; }
        public string unit { get; set; }


    }

}
