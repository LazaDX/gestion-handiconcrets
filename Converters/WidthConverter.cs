using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace gestion_concrets.Converters
{
    public class WidthConverter : IValueConverter
    {
        public static readonly WidthConverter Instance = new WidthConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double width)
            {
                return width - 20; // Ajustement pour les marges (10 de chaque côté)
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
