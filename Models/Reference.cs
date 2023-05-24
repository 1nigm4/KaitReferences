using System;

namespace KaitReferences.Models
{
    public class Reference
    {
        public DateTime Date { get; set; }
        public string Type { get; set; }
        public int Count { get; set; }
        public string Assignment { get; set; }
        public string Period { get; set; }
        public string Form { get; set; }
        public string Note { get; set; }
        public string Status { get; set; }
        public ReferenceType ReferenceType { get; set; }
    }
}
