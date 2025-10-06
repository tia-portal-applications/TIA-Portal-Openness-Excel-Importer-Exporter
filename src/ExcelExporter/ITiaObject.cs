using Siemens.Engineering;
using Siemens.Engineering.HmiUnified;
using System.Collections.Generic;

namespace ExcelExporter
{
    interface ITiaObject
    {
        Dictionary<string, List<Dictionary<string, object>>> Export(HmiSoftware hmiSoftware);
        void Import(HmiSoftware hmiSoftware, List<object> data);



    }
}
