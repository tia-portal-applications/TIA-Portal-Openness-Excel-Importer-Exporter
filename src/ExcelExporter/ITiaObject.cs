using Siemens.Engineering;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.UI.Screens;
using System.Collections.Generic;

namespace ExcelExporter
{
    interface ITiaObject
    {
        Dictionary<string, List<Dictionary<string, object>>> Export(IEnumerable<HmiScreen> allScreens);
        void Import(HmiSoftware hmiSoftware, List<object> data, IEnumerable<HmiScreen> allScreens);



    }
}
