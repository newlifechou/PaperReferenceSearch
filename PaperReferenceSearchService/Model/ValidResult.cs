using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearchService.Model
{
    public class ValidResult
    {
        public ValidResult()
        {
            IsValid = true;
            ValidInformation = "";
        }
        public ValidResult(bool isValid,string validInformation)
        {
            IsValid = isValid;
            ValidInformation = validInformation;
        }
        public bool IsValid { get; set; }
        public string ValidInformation { get; set; }
    }
}
