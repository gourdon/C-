using System;
using System.Collections.Generic;
using System.Text;

namespace Model
{
    public enum Status : byte
    {
        /// <summary>
        /// Defect has been opened, but not verified as reproducible or an issue.
        /// </summary>
        Created,
        /// <summary>
        /// Defect has been verified as an issue requiring work.
        /// </summary>
        Accepted,
        /// <summary>
        /// Defect has been fixed in code, but not verified other than through developer testing.
        /// </summary>
        Fixed,
        /// <summary>
        /// Defect was fixed, but has now been reopened due to failing verification.
        /// </summary>
        Reopened,
        /// <summary>
        /// Defect has been fixed and tested; the fix is satisfactory.
        /// </summary>
        Closed,
    }
}
