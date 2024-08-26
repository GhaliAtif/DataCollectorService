using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCollectorService
{
    public class DataEntry2
    {
        public DateTime Date { get; set; }
        public TimeSpan Time { get; set; }
        public string Element { get; set; }
        public string SampleID { get; set; }
        public double Concentration { get; set; }
        public string Unit { get; set; }
    }
}
