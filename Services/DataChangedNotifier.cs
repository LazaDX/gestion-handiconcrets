using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;


namespace gestion_concrets.Services
{
    public class DataChangedNotifier
    {
        public static event EventHandler DataChanged;

        public static void NotifyDataChanged()
        {
            DataChanged?.Invoke(null, EventArgs.Empty);
        }
    }
}
