﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocParser
{
    class Question
    {
        public string Content { get; set; }
        public string Type { get; set; }

        public List<Answer> Answers { get; set; }
    }
}
