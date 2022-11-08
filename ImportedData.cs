using System;
using System.Collections.Generic;
using System.Text;

namespace LV_Metrics_Parser
{
    public class ImportedData
    {
        public string Minas { get; set; }
        public string Kata { get; set; }
        public string Sys { get; set; }
        public string Domi { get; set; }
        public string Xrisi { get; set; }
        public string Timologio { get; set; }
        public string ArLogar { get; set; }
        public string KWH_H { get; set; }
        public string KWH_N { get; set; }
        public string KWH_T { get; set; }
        public string Pagio { get; set; }
        public string ElXreosi { get; set; }
        public string Energia { get; set; }
        public string Isxys { get; set; }
        public string DEnergias { get; set; }
        public string CO2 { get; set; }
        public string EkptEt { get; set; }
        public string EkptOgk { get; set; }
        public string AntEkpt { get; set; }
        public string EkptTY { get; set; }
        public string TK_KY { get; set; }
        public string Mig { get; set; }
        public string EkptPag { get; set; }
        public string NeaEkptPagiou { get; set; }
        public string EkptPoso { get; set; }
        public string Ekpt8 { get; set; }
        public string EkptosiGP { get; set; }
        public string Ekpt10 { get; set; }
        public string Ekpt15 { get; set; }
        public string Pro { get; set; }
        public string EkpOgkDim { get; set; }
        public string Ekpt2 { get; set; }
        public string EkptEnXT { get; set; }
        public string Kouponi { get; set; }
        public string EkptEbill { get; set; }
        public string GreenPass { get; set; }
        public string EkGreenPass { get; set; }
        public string XXS { get; set; }
        public string XXD { get; set; }
        public string YKO { get; set; }
        public string LX { get; set; }
        public string ETMEAR { get; set; }
        public string EFK { get; set; }
        public string DETE { get; set; }
        public string FPA { get; set; }
        public string EpidPOT { get; set; }
        public string EpidKOT { get; set; }
        public string EpidKOT50 { get; set; }
        public string XR_1 { get; set; }


        public ImportedData ImportedDataClone()
        {
            return (ImportedData)this.MemberwiseClone();
        }
    }
}
