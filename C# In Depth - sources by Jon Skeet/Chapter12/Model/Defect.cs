using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Model
{
    public partial class Defect
    {
        int count=1;

        partial void OnCreated()
        {
            ID=count++;
        }

        public override string ToString()
        {
            return string.Format("{0,2}: {1}\r\n    ({2:d}-{3:d}, {4}/{5}, {6} -> {7})",
                ID, Summary, Created, LastModified, Severity, Status, CreatedBy.Name,
                AssignedTo == null ? "n/a" : AssignedTo.Name);
        }
    }
}
