using Siemens.Engineering;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HW.Features;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelExporter
{
    class Program
    {
        static string opennessDll;
        static public bool verbose = false;
        static private Assembly DomainAssemblyResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf("Siemens.Engineering,");

            if (index != -1 || args.Name == "Siemens.Engineering")
            {
                return Assembly.LoadFrom(opennessDll);
            }

            return null;
        }
        static void Main(string[] args)
        {
            // 0. Load latest TIA Openness DLL dynamically, so it works also in all further versions
            string[] dirs = Directory.GetDirectories(@"C:\Program Files\Siemens\Automation\", "Portal *");
            // opennessDll = @"C:\Program Files\Siemens\Automation\Portal V16\PublicAPI\V16\Siemens.Engineering.dll";
            var latestVersionDirectory = dirs[dirs.Length - 1];
            string[] opennessDirs = Directory.GetDirectories(latestVersionDirectory + "\\PublicAPI", "V*");
            var latestOpennessDir = opennessDirs.Last(x => !x.ToLower().EndsWith("addin"));
            opennessDll = latestOpennessDir + "\\Siemens.Engineering.dll";
            AppDomain.CurrentDomain.AssemblyResolve += DomainAssemblyResolver;

            // TiaPortal reference must be in a seperate function to lazyload the assembly
            Work(args);
            Console.WriteLine("Export finished");
        }
        static void Work(string[] args)
        {
            string screenName = "";
            string runTimeName = "";
            bool setProperties = false; 
            List<string> definedAttributes = null;
            while (args.Length == 0)
            {
                Console.WriteLine("Please define a screen name (or set -all to export all screens):");
                Console.WriteLine("If you want to export only defined attributes, please start this tool again with arguments like this: ");
                Console.WriteLine("HMI_Panel/Screen_1 Left Top Authorization");
                Console.WriteLine("You can also add the verbose flag like this for more output: HMI_Panel/Screen_1 Left Top Authorization --verbose");
                args = Console.ReadLine().Split(' ');
            }
            
                string firstInput = args[0];
                if (firstInput.Contains('/'))
                {
                    runTimeName = firstInput.Split('/')[0];
                    screenName = firstInput.Split('/')[1];
                }
                else
                {
                    screenName = args[0];
                }

                if (args.Length > 1)
                {
                    definedAttributes = args.ToList();
                    if (definedAttributes.Contains("--verbose"))
                    {
                        verbose = true;
                        definedAttributes.Remove("--verbose");
                    }
                    definedAttributes.RemoveAt(0); // remove first input (Runtime/screenname)
                    setProperties = true; 
                }
            
            YAMLConverter converter = new YAMLConverter();
            var defaultValues = (converter.Read(Directory.GetCurrentDirectory() + "\\DefaultScreen.yml")[0] as Dictionary<object, object>).FirstOrDefault().Value as List<object>;

            Dictionary<string, List<Dictionary<string, object>>> exportedValues = new Dictionary<string, List<Dictionary<string, object>>>();
            // 1. Connect to TIA Portal and find screen
            TiaPortal tiaPortal = null;
            try
            {
                // get UnifiedRuntime
                var hmiSoftwares = GetHmiSoftwares(ref tiaPortal, runTimeName);
                HmiSoftware software = hmiSoftwares.FirstOrDefault(); //find runtimeName
                if (software == null)
                {
                    throw new Exception("No WinCC Unified software found. Please add a WinCC Unified device and run this app again!");
                }
                Screen screen = new Screen(screenName, definedAttributes);
                exportedValues = screen.Export(software);
            }
            finally
            {
                tiaPortal.Dispose();
            }
            CreateExport(exportedValues, defaultValues, setProperties);
            Console.WriteLine("Export done");
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
                    Console.WriteLine("Cannot find default type with name: " + type);
                }
                else
                {
                    if (verbose) Console.WriteLine(defaultObject.ToString());
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
                        if (attribute.Value == null)
                        {
                            differencesScreenItem.Add(attribute.Key, attribute.Value);
                        }
                        else if (string.Compare(attribute.Value.ToString(), defaultValue.ToString(), true) != 0)
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

        static IEnumerable<HmiSoftware> GetHmiSoftwares(ref TiaPortal tiaPortal, string deviceName)
        {
            var tiaProcesses = TiaPortal.GetProcesses();
            if (tiaProcesses.Count == 0)
            {
                new Exception("No TIA Portal instance is running. Please start TIA Portal and open a project with a WinCC Unified device and run this app again!");
            }
            var process = tiaProcesses[0];
            tiaPortal = process.Attach();
            Console.WriteLine("Attached to TIA Portal process id: " + process.Id);
            if (tiaPortal.Projects.Count == 0)
            {
                new Exception("No TIA Portal project is open. Please open a project with a WinCC Unified device and run this app again!");
            }
            Project tiaPortalProject = tiaPortal.Projects.First();
            Console.WriteLine("Attached to TIA Portal project with name: " + tiaPortalProject.Name);
            var software = from device in tiaPortalProject.Devices
                           from deviceItem in device.DeviceItems
                           let softwareContainer = deviceItem.GetService<SoftwareContainer>()
                           where softwareContainer?.Software is HmiSoftware && (string.IsNullOrWhiteSpace(deviceName) || device.Name == deviceName)
                           select softwareContainer.Software as HmiSoftware;
            if (software == null)
            {
                software = from device in tiaPortalProject.Devices
                               from deviceItem in device.DeviceItems
                               let softwareContainer = deviceItem.GetService<SoftwareContainer>()
                               where softwareContainer?.Software is HmiSoftware
                               select softwareContainer.Software as HmiSoftware;
            }
            return software;
                
        }
    }
}
