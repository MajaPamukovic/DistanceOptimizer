using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace DistanceOptimizer
{
    public class OfficeLocation
    {
        public OfficeLocation()
        {
            AverageDuration = 0;
            MinimalDuration = Double.PositiveInfinity;
            MaximalDuration = 0;
            NonZeroDurations = 0;
        }

        public string Address { get; set; }

        public string Name { get; set; }

        [ScriptIgnore]
        public double AverageDuration { get; set; }

        [ScriptIgnore]
        public double MinimalDuration { get; set; }

        [ScriptIgnore]
        public double MaximalDuration { get; set; }

        [ScriptIgnore]
        public int NonZeroDurations { get; set; }

        [ScriptIgnore]
        public double CalculateAverageDuration
        {
            get
            {
                return AverageDuration / NonZeroDurations;
            }
        }
    }
}
