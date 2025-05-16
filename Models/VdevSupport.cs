using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace gestion_concrets.Models
{
    public class VdevSupport
    {
        public int Id { get; set; }
        public int IdBPerson { get; set; }  // Clé étrangère vers BPerson

        public string V1 { get; set; }
        public string V2 { get; set; }
        public string V3 { get; set; }
        public string V41 { get; set; }
        public string V42 { get; set; }
        public string V51 { get; set; }
        public string V52 { get; set; }
        public string V53 { get; set; }
    }
}
