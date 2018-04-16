using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace PaperReferenceSearch
{
    public class ValidConverter : IValueConverter
    {
        public ValidConverter()
        {
            Good = "参与";
            Bad = "跳过";
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
