using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace gestion_concrets.Models
{
    public class BPerson
    {
        public int Id { get; set; }
        public string B1 { get; set; }
        public string B2 { get; set; }
        private DateTime? _dateTime;
        public DateTime? DateTime // Pour le DatePicker
        {
            get => _dateTime;
            set
            {
                _dateTime = value;
                B2 = _dateTime?.ToString("dd/MM/yyyy") ?? "Non définie";
            }
        }
        public string B3 { get; set; }
        public string Adress { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string A1 { get; set; } // Date de recensement
        private DateTime? _dateAndTime;
        public DateTime? DateAndTime
        {
            get => _dateAndTime;
            set
            {
                _dateAndTime = value;
                B2 = _dateAndTime?.ToString("dd/MM/yyyy") ?? "Non définie";
            }
        }
        public string B4 { get; set; }
        public string B42 { get; set; }
        public string B51 { get; set; }
        public string B52 { get; set; }
        public string B6 { get; set; }
        public string B61 { get; set; }
        public string B7 { get; set; }
        public string B71 { get; set; }
        public string B72 { get; set; }
        public string B73 { get; set; }
        public string B74 { get; set; }

        public string B8 { get; set; }
    }
}
