using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace DistanceOptimizer
{
    public class ValueNode
    {
        public string text { get; set; }
        public double value { get; set; }
    }
    
    public class Element
    {
        public ValueNode distance { get; set; }
        public ValueNode duration { get; set; }
        public ValueNode duration_in_traffic { get; set; }
        public string status { get; set; }
    }

    public class Row
    {
        public List<Element> elements { get; set; }
    }

    public class RouteResponse
    {
        public List<string> destination_addresses { get; set; }
        public List<string> origin_addresses { get; set; }
        public List<Row> rows { get; set; }

        [ScriptIgnore]
        public double DurationInTraffic
        {
            get
            {
                if (rows == null || rows.Count() == 0 || rows.First() == null || rows.First().elements == null || rows.First().elements.First() == null || rows.First().elements.First().duration_in_traffic == null)
                    return 0;

                return rows.First().elements.First().duration_in_traffic.value;
            }
        }

        [ScriptIgnore]
        public double Duration
        {
            get
            {
                if (rows == null || rows.Count() == 0 || rows.First() == null || rows.First().elements == null || rows.First().elements.First() == null || rows.First().elements.First().duration == null)
                    return 0;

                return rows.First().elements.First().duration.value;
            }
        }

        [ScriptIgnore]
        public double GetBestDurationInfo
        {
            get
            {
                return DurationInTraffic == 0 ? Duration : DurationInTraffic;
            }
        }

        [ScriptIgnore]
        public double Distance
        {
            get
            {
                if (rows == null || rows.Count() == 0 || rows.First() == null || rows.First().elements == null || rows.First().elements.First() == null || rows.First().elements.First().distance == null)
                    return -1;

                return rows.First().elements.First().distance.value;
            }
        }

        [ScriptIgnore]
        public string Status
        {
            get
            {
                if (rows == null || rows.Count() == 0 || rows.First() == null || rows.First().elements == null || rows.First().elements.First() == null)
                    return "";
                string result = rows.First().elements.First().status;
                if (result == "OVER_QUERY_LIMIT")
                    throw new UnauthorizedAccessException("Google API request quota exceeded");
                return result;
            }
        }
    }
}
