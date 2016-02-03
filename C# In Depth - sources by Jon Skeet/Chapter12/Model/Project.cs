using System;
using System.Collections.Generic;
using System.Text;

namespace Model
{
    public partial class Project
    {
        public override string ToString()
        {
            return string.Format("Project: {0}", Name);
        }
    }
}
