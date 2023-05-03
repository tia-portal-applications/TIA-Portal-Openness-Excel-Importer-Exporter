using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExporter
{
    interface IConverter
    {
        void Write(Dictionary<string, List<Dictionary<string, object>>> data, string path);
        List<object> Read(string path);
    }
}
