using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

namespace InsightsaddinWebApp.Models
{
    public class PartnerAccount
    {
        [JsonProperty(PropertyName = "pbe")]
        public string Pbe { get; set; }

        [JsonProperty(PropertyName = "website")]
        public string Website { get; set; }

        [JsonProperty(PropertyName = "crm")]
        public string Crm { get; set; }

        [JsonProperty(PropertyName = "stage")]
        public string Stage { get; set; }

        [JsonProperty(PropertyName = "engagementType")]
        public string EngagementType { get; set; }

        [JsonProperty(PropertyName = "date")]
        public string Date { get; set; }

        [JsonProperty(PropertyName = "reason")]
        public string Reason { get; set; }

        [JsonProperty(PropertyName = "location")]
        public string Location { get; set; }

        [JsonProperty(PropertyName = "meeting")]
        public string Meeting { get; set; }

        [JsonProperty(PropertyName = "industry")]
        public string Industry { get; set; }

        [JsonProperty(PropertyName = "cloudStatus")]
        public string CloudStatus { get; set; }

        [JsonProperty(PropertyName = "cloudProvider")]
        public string CloudProvider{ get; set; }

        [JsonProperty(PropertyName = "consumption")]
        public string Consumption { get; set; }

        [JsonProperty(PropertyName = "workLoads")]
        public string WorkLoads { get; set; }
    }
}