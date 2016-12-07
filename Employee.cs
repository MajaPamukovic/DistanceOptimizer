using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace DistanceOptimizer
{
    public class Employee
    {
        private static string GetEnumDescription(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(
                typeof(DescriptionAttribute),
                false);

            if (attributes != null &&
                attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }

        public Employee()
        {
            // ???
        }

        public string Name { get; set; }

        public string Address { get; set; }

        [ScriptIgnore]
        public TransitMode TransitMode { get; set; }

        public string TransitModeString
        {
            get
            {
                return GetEnumDescription(TransitMode);
            }
            set
            {
                TransitMode transitModeVal = TransitMode.Driving;
                Enum.TryParse<TransitMode>(value, out transitModeVal);
                TransitMode = transitModeVal;
            }
        }
    }
}
