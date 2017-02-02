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
        driving,
        [Description("transit")]
        transit,
        [Description("bicycling")]
        bicycling
    }
}
