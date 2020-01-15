using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearchService.Model
{
    public class Author
    {
        public Author()
        {
            Name = "";
            IsMatched = false;
        }
        public string Name { get; set; }
        public bool IsMatched { get; set; }
    }
}
