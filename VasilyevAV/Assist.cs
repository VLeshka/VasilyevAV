using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VasilyevAV
{
    static class Assist
    {
        string[] fileExtChecks = { ".xml", ".pke" };
        char[] xmlTrim = { ' ', '/', '<', '>' };

        enum t_pke_cxema { cxema1, cxema2, error };

        struct t_RM3
        {
            public string ver;
            public string uid;
        }

        struct t_Check
        {
            public t_RM3 rmz;
            public ListViewItem itemCheck;
            public Dictionary<int, ListViewItem> itemsCheckDetails; //№пп, ListViewItem
            public t_pke_cxema pke_cxema()
            {
                const int cxemaIndex = 3;
                try
                {
                    if (itemCheck.SubItems[cxemaIndex].Text == "1")
                        return (t_pke_cxema.cxema1);
                    if (itemCheck.SubItems[cxemaIndex].Text == "2")
                        return (t_pke_cxema.cxema2);
                    return (t_pke_cxema.error);
                }
                catch
                {
                    return (t_pke_cxema.error);
                }
            }
        }

        Dictionary<string, t_Check> checks = new Dictionary<string, t_Check>(); //UID, даные проверки
    }
}
