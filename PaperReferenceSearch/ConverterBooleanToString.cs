using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace PaperReferenceSearch
{
    public class ConverterBooleanToString : IValueConverter
    {
        public ConverterBooleanToString()
        {
            Good = "PASS";
            Bad =  "FAIL";
        }
        public string Good { get; set; }
        public string Bad { get; set; }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool isValid = (bool)value;
            return isValid ? Good : Bad;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
