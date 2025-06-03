using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace gestion_concrets.Models
{
    public class Alocalisation
    {
        public int Id { get; set; }
        public int IdBPerson { get; set; } 
        public string A1 { get; set; }
        private DateTime? _dateTime;
        public DateTime? DateTime // Pour le DatePicker
        {
            get => _dateTime;
            set
            {
                _dateTime = value;
                A1 = _dateTime?.ToString("dd/MM/yyyy") ?? "Non définie";
            }
        }
        public string A2 { get; set; }
        public string A3 { get; set; }
        public string A4 { get; set; }

    }
}
