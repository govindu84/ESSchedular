using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ESScheduler.Models
{
    public class ESAPP_Input
    {
        public string OrderCode { get; set; }
        public dynamic RevPrice { get; set; }
        public dynamic RetailPrice { get; set; }
        public dynamic SalePrice { get; set; }
        public dynamic TotalDiscount { get; set; }
      
        
    }
    public static class LightweightProfileData
    {

        public static string Country { get; set; } = "us";
        public static string Region { get; set; } = "us";

        public static string Language { get; set; } = "en";

        public static string Segment { get; set; } = "dhs";

        public static string CustomerSet { get; set; } = "19";

    }


}
