using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleJob
{
    public class MerchantModel
    {
        public string ISOName { get; set; }
        public string AgentSalesOfficeName { get; set; }
        public string AgentSalesOffice { get; set; }
        public string MerchantMID { get; set; }
        public string MerchantName { get; set; }
        public string DateBoardedOpen { get; set; }
        public string DateClosed { get; set; }
        public string ProgramType { get; set; }
        public string Platform { get; set; }
        public string TransCount { get; set; }
        public string SalesVolume { get; set; }
        public string SalesChannel { get; set; }
        public string Status { get; set; }
    }

    /*public class ModelClassMap : ClassMap<MerchantModel>
    {
        public ModelClassMap()
        {
            Map(m => m.ISOName).Name("ISO Name");
            Map(m => m.AgentSalesOfficeName).Name("Agent Sales Office Name");
            Map(m => m.AgentSalesOffice).Name("Agent Sales Office #");
            Map(m => m.MerchantMID).Name("Merchant # (MID)");
            Map(m => m.MerchantName).Name("Merchant Name");
            Map(m => m.DateBoardedOpen).Name("Date Boarded/Open");
            Map(m => m.DateClosed).Name("Date Closed");
            Map(m => m.ProgramType).Name("Program Type");
            Map(m => m.Platform).Name("Platform");
            Map(m => m.TransCount).Name("Trans Count");
            Map(m => m.SalesVolume).Name("Sales Volume");
            Map(m => m.SalesChannel).Name("Sales Channel");
            Map(m => m.Status).Name("Status");
        }
    }*/


}
