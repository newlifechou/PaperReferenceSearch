﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    public class JobUnit
    {
        public JobUnit()
        {
            Master = new Paper();
            References = new List<Paper>();
        }
        public Paper Master { get; set; }
        public List<Paper> References { get; set; }

        public int ReferenceCount
        {
            get
            {
                return References.Count();
            }
        }

        public int SelfReferenceCount
        {
            get
            {
                return References.Where(i => i.ReferenceType == PaperReferenceType.Self).Count();
            }
        }

        public int OtherReferenceCount
        {
            get
            {
                return References.Where(i => i.ReferenceType == PaperReferenceType.Other).Count();
            }
        }

        public int UnSetReferenceCount
        {
            get
            {
                return References.Where(i => i.ReferenceType == PaperReferenceType.UnSet).Count();
            }
        }


    }
}
