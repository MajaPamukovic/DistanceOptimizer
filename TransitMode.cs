using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DistanceOptimizer
{
    public enum TransitMode
    {
        [Description("driving")]
        Driving,
        [Description("transit")]
        Transit,
        [Description("bicycling")]
        Bicycling
    }
}
