using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Tag;
using UnifiedOpennessLibrary;

namespace ExcelExporter
{
    class Program
    {
        public static UnifiedOpennessConnector unifiedData = null;
        static void Main(string[] args)
        {
            using (var unifiedData = new UnifiedOpennessConnector("V19", args, new List<CmdArgument>() { new CmdArgument()
            {
                Default = "", Required = false, OptionToSet = "DefinedAttributes", OptionLong = "--definedattributes", OptionShort = "-da", HelpText = "If you want to export only defined attributes, add a list seperated by semicolon, e.g. Left;Top;Authorization"
            } }, "ExcelExporter"))
            {
                Program.unifiedData = unifiedData;
                Work();
            }
            unifiedData.Log("Export finished");
        }
        static void Work()
        {
            bool setProperties = false; 
            List<string> definedAttributes = null;
            if (!string.IsNullOrEmpty(unifiedData.CmdArgs["DefinedAttributes"]))
            {
                definedAttributes = unifiedData.CmdArgs["DefinedAttributes"].Split(';').ToList();
                setProperties = true; 
            }
            
            YAMLConverter converter = new YAMLConverter();
            var defaultValues = (converter.Read(Directory.GetCurrentDirectory() + "\\DefaultScreen.yml")[0] as Dictionary<object, object>).FirstOrDefault().Value as List<object>;

            Dictionary<string, List<Dictionary<string, object>>> exportedValues = new Dictionary<string, List<Dictionary<string, object>>>();
           
            Screen screen = new Screen(definedAttributes);
            exportedValues = screen.Export(unifiedData.Screens);
            CreateExport(exportedValues, defaultValues, setProperties);
            unifiedData.Log("Export done");
        }

        private static void CreateExport(Dictionary<string, List<Dictionary<string, object>>> exportedValues, List<object> defaultValues, bool setProperties)
        {
            foreach (var screenDict in exportedValues)
            {

                var differences = GetDifferences(screenDict.Value, defaultValues, setProperties);

                string filename = Directory.GetCurrentDirectory() + "\\" + screenDict.Key + ".xlsx";
                Microsoft.Office.Interop.Excel.Application xlApp = null;
                Microsoft.Office.Interop.Excel.Workbook workbook = null;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    workbook = xlApp.Workbooks.Add();
                    worksheet = workbook.Worksheets[1];

                        writeToCells(ref worksheet, differences);

                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }
                    object misValue = System.Reflection.Missing.Value;
                    workbook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
                    misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


                    workbook.Close(true, misValue, misValue);
                    xlApp.Quit();
                }
                finally
                {
                    Marshal.ReleaseComObject(xlApp);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(worksheet);
                    xlApp = null;
                    workbook = null;
                    worksheet = null;
                }
            }
        }

        private static void writeToCells(ref Worksheet worksheet, List<Dictionary<string, object>> differences)
        {
            worksheet.Cells[1, 1] = "Type";
            worksheet.Cells[1, 2] = "Name";
            Dictionary<string, int> columnContentToIndex = new Dictionary<string, int>();
            columnContentToIndex.Add("Type", 1);
            columnContentToIndex.Add("Name", 2);
            for (int i = 0; i < differences.Count; i++)
            {
                writePropertiesToCells(ref worksheet, i + 2, differences[i], ref columnContentToIndex, "");
            }
        }

        private static void writePropertiesToCells(ref Worksheet worksheet, int rowIndex, Dictionary<string, object> properties, ref Dictionary<string, int> columnContentToIndex, string parentName)
        {
            foreach (var item in properties)
            {
                if (item.Value is Dictionary<string, object>)
                {
                    writePropertiesToCells(ref worksheet, rowIndex, item.Value as Dictionary<string, object>, ref columnContentToIndex, parentName + item.Key + ".");
                }
                else if (item.Value is List<object>)
                {
                    var list = (item.Value as List<object>);
                    for (int i = 0; i < list.Count; i++)
                    {
                        writePropertiesToCells(ref worksheet, rowIndex, list[i] as Dictionary<string, object>, ref columnContentToIndex, parentName + item.Key + "[" + i + "].");
                    }
                }
                else if (item.Value is ValueConverter)
                {
                    // TODO: implement to unpack the values. ValueConverter can be Range, Bitmask, Singlebit
                }
                else
                {
                    int columnIndex = -1;
                    if (!columnContentToIndex.TryGetValue(parentName + item.Key, out columnIndex))
                    {
                        columnIndex = columnContentToIndex.Count + 1; // worksheet starts at index 1
                        columnContentToIndex.Add(parentName + item.Key, columnIndex);
                        worksheet.Cells[1, columnIndex] = parentName + item.Key;
                    }
                    var value = item.Value;
                    if (item.Value is List<string>)
                    {
                        value = string.Join(",", item.Value as List<string>);
                    }
                    else if (value is bool)
                    {
                        // add a space at the end to make sure the english words for true and false will be used. Otherwise Excel will use the local language and it cannot be imported on another PC with different language.
                        value = ((bool)value == true) ? "True " : "False ";
                    }
                    worksheet.Cells[rowIndex, columnIndex] = value;
                }
            }
        }

        private static List<Dictionary<string, object>> GetDifferences(List<Dictionary<string, object>> screenItems, List<object> defaultValues, bool setProperties)
        {
            List<Dictionary<string, object>> differences = new List<Dictionary<string, object>>();
            foreach (Dictionary<string, object> attributes in screenItems)
            {
                string type = attributes["Type"].ToString();
                //if (type == "HmiScreen")
                //{
                //    continue; // screens will not be handled
                //}
                var defaultObject = defaultValues.Find(x =>
                {
                    object typeName = null;
                    (x as Dictionary<object, object>).TryGetValue("Type", out typeName);
                    return typeName.ToString() == type;
                }) as Dictionary<object, object>;
                if (defaultObject == null)
                {
                    unifiedData.Log("Cannot find default type with name: " + type, LogLevel.Warning);
                }
                else
                {
                    unifiedData.Log(defaultObject.ToString(), LogLevel.Debug);
                    var differencesScreenItem = new Dictionary<string, object>();
                        GetDifferencesScreenItem(attributes, defaultObject, ref differencesScreenItem);
                    differencesScreenItem.Add("Type", type); // type is always needed to create the object again
                    if (setProperties)
                    {
                        differences.Add(attributes);
                    }
                    else
                    {
                        differences.Add(differencesScreenItem);
                    }
                }
            }
            return differences;
        }

        private static void GetDifferencesScreenItem(Dictionary<string, object> attributes, Dictionary<object, object> defaultObject, ref Dictionary<string, object> differencesScreenItem)
        {
            foreach (var attribute in attributes)
            {
                if (attribute.Value is List<object>)
                {
                    var list = attribute.Value as List<object>;
                    object deeperObj = null;
                    var deeperList = new List<object>();
                    if (defaultObject.TryGetValue(attribute.Key, out deeperObj))
                    {
                        deeperList = deeperObj as List<object>;
                    }
                    var newDiffList = new List<object>();
                    for (int i = 0; i < list.Count; i++)
                    {
                        var newDiffValue = new Dictionary<string, object>();
                        var deeperDict = new Dictionary<object, object>();
                        if (i < deeperList.Count)
                        {
                            deeperDict = deeperList[i] as Dictionary<object, object>;
                        }
                        GetDifferencesScreenItem(list[i] as Dictionary<string, object>, deeperDict, ref newDiffValue);
                        if (newDiffValue.Count > 0)
                        {
                            newDiffList.Add(newDiffValue);
                        }
                    }
                    if (newDiffList.Count > 0)
                    {
                        differencesScreenItem.Add(attribute.Key, newDiffList);
                    }
                }
                else if (attribute.Value is Dictionary<string, object>)
                {
                    object deeperObj = null;
                    var deeperDict = new Dictionary<object, object>();
                    if (defaultObject.TryGetValue(attribute.Key, out deeperObj))
                    {
                        deeperDict = deeperObj as Dictionary<object, object>;
                    }
                    var newDiffValue = new Dictionary<string, object>();
                    GetDifferencesScreenItem(attribute.Value as Dictionary<string, object>, deeperDict, ref newDiffValue);
                    if (newDiffValue.Count > 0)
                    {
                        differencesScreenItem.Add(attribute.Key, newDiffValue);
                    }
                }
                else
                {
                    object defaultValue = null;
                    if (defaultObject.TryGetValue(attribute.Key, out defaultValue) && defaultValue != null)
                    {
                        if (string.Compare(attribute.Value.ToString(), defaultValue.ToString(), true) != 0)
                        {
                            differencesScreenItem.Add(attribute.Key, attribute.Value);
                        }
                    }
                    else
                    {
                        differencesScreenItem.Add(attribute.Key, attribute.Value);
                    }
                }
            }
        }
    }
}
