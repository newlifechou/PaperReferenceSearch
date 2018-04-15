using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearch.Model
{
    public class DataFile
    {
        public string Name { get; set; }
        public string FullName { get; set; }
        public bool ValidState { get; set; }
        public string ValidInformation { get; set; }
    }
}
