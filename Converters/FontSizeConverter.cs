using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace gestion_concrets.Converters
{
    internal class FontSizeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (parameter == null) return 12.0; // Valeur par défaut si parameter est null
            if (!double.TryParse(parameter.ToString(), out double baseSize))
                return 12.0; // Valeur par défaut en cas d'erreur de parsing
            double scaleFactor = GetScaleFactor();
            return baseSize * scaleFactor;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        private double GetScaleFactor()
        {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            if (screenWidth <= 1366) return 0.8; // Petits écrans
            if (screenWidth <= 1920) return 1.0; // Écrans standards
            return 1.2; // Écrans HD ou plus grands
        }
    }

    // Convertisseur pour MinWidth
    public class MinWidthConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (parameter == null) return 150.0; // Valeur par défaut si parameter est null
            if (!double.TryParse(parameter.ToString(), out double baseWidth))
                return 150.0; // Valeur par défaut en cas d'erreur de parsing
            double scaleFactor = GetScaleFactor();
            return baseWidth * scaleFactor;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        private double GetScaleFactor()
        {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            if (screenWidth <= 1366) return 0.8; // Petits écrans
            if (screenWidth <= 1920) return 1.0; // Écrans standards
            return 1.2; // Écrans HD ou plus grands
        }
    }

    // Convertisseur pour MinHeight
    public class MinHeightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (parameter == null) return 25.0; // Valeur par défaut si parameter est null
            if (!double.TryParse(parameter.ToString(), out double baseHeight))
                return 25.0; // Valeur par défaut en cas d'erreur de parsing
            double scaleFactor = GetScaleFactor();
            return baseHeight * scaleFactor;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        private double GetScaleFactor()
        {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            if (screenWidth <= 1366) return 0.8; // Petits écrans
            if (screenWidth <= 1920) return 1.0; // Écrans standards
            return 1.2; // Écrans HD ou plus grands
        }
    }

    // Convertisseur pour Margin
    public class ResolutionMarginConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            if (screenWidth <= 1366) return new Thickness(3); // Petits écrans
            if (screenWidth <= 1920) return new Thickness(5); // Écrans standards
            return new Thickness(7); // Écrans HD ou plus grands
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

