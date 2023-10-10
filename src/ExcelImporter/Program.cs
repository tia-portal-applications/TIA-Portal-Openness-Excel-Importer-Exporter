using Siemens.Engineering;
using Siemens.Engineering.HmiUnified;
using Siemens.Engineering.HmiUnified.UI.Base;
using Siemens.Engineering.HmiUnified.UI.Screens;
using Siemens.Engineering.HmiUnified.UI.Shapes;
using Siemens.Engineering.HW.Features;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Siemens.Engineering.HmiUnified.UI.Dynamization;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Script;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Flashing;
using Siemens.Engineering.HmiUnified.UI;
using Siemens.Engineering.HmiUnified.UI.Controls;
using Siemens.Engineering.HmiUnified.UI.Widgets;
using Siemens.Engineering.HmiUnified.UI.ScreenGroup;

namespace ExcelImporter
{
    public static class Globals
    {
        public static String oldPropertyName = ""; // Modifiable
        public static ScriptDynamization scriptTemp= null;
    }
    class Program
    {
        static string opennessDll;
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
            Console.WriteLine("Import finished");
        }

        private static IEnumerable<HmiScreen> GetScreens(HmiSoftware sw)
        {
            var allScreens = sw.Screens.ToList();
            allScreens.AddRange(ParseGroups(sw.ScreenGroups));
            return allScreens;
        }

        private static IEnumerable<HmiScreen> ParseGroups(HmiScreenGroupComposition parentGroups)
        {
            foreach (var group in parentGroups)
            {
                foreach (var screen in group.Screens)
                {
                    yield return screen;
                }
                foreach (var screen in ParseGroups(group.Groups))
                {
                    yield return screen;
                }
            }
        }


        static void Work(string[] args) {

            // 1. Connect to TIA Portal and find screen
            TiaPortal tiaPortal = null;
            try
            {
                // get UnifiedRuntime
                var hmiSoftwares = GetHmiSoftwares(ref tiaPortal);
                HmiSoftware software = hmiSoftwares.FirstOrDefault();
                if (software == null)
                {
                    new Exception("No WinCC Unified software found. Please add a WinCC Unified device and run this app again!");
                }

                //2.Open Excel sheet
                string[] files = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx");
                var allScreens = GetScreens(software);
                //iterate over every file
                foreach (string file in files)
                {
                    string filename = Path.GetFileName(file);
                    Microsoft.Office.Interop.Excel.Application xlApp = null;
                    Microsoft.Office.Interop.Excel.Workbook workbook = null;
                    Microsoft.Office.Interop.Excel.Range range = null;
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

                    try
                    {
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        workbook = xlApp.Workbooks.Open((Directory.GetCurrentDirectory() + "\\" + filename));
                        worksheet = workbook.Worksheets[1];
                        range = workbook.ActiveSheet.UsedRange;


                        if (!File.Exists(file))
                        {
                            Console.WriteLine("Cannot find a file with name '" + filename + "' in path '" + Directory.GetCurrentDirectory() + "'");
                            Console.WriteLine("Please place a file with name '" + filename + "' next to the app!");
                            return;
                        }



                        // get screen filename=foo.xlsx
                        var screenName = filename.Split('.')[0];
                        var screen = allScreens.FirstOrDefault(s => s.Name == screenName);
                        if (screen == null)
                        {
                            screen = software.Screens.Create(screenName);
                            Console.WriteLine("New screen with name '" + screenName + "' added.");
                        }
                        else
                        {
                            Console.WriteLine("Found screen with name '" + screenName + "'.");
                        }
                        // 3. Read Excel file and add elements to TIA portal
                        int tableColumn = 1;
                        int tableRow = 2;

                        while (true)
                        {
                            if (worksheet.Cells[1][tableRow].value2 == null)
                            {
                                break; // end of file
                            }
                            Dictionary<string, object> propertyNameValues = new Dictionary<string, object>();
                            while (true)
                            {
                                //Check if cell is empty and add to dictonary
                                if (worksheet.Cells[tableColumn][tableRow].value2 == null)
                                {
                                    propertyNameValues.Add(worksheet.Cells[tableColumn][1].value2, "");
                                }
                                else
                                {
                                    propertyNameValues.Add(worksheet.Cells[tableColumn][1].value2, worksheet.Cells[tableColumn][tableRow].value2);
                                }

                                tableColumn++;
                                //check if all attributes are read in and Create the Screen Item 
                                if (worksheet.Cells[tableColumn][1].value2 == null)
                                {
                                    tableColumn = 1;
                                    break;
                                }
                            }
                            CreateScreenItem(screen, propertyNameValues);
                            Console.WriteLine();
                            tableRow++;
                        }

                        for (int tableRow_ = 2; worksheet.Cells[1][tableRow_].value2 != null; tableRow_++)
                        {
                            for (int tableColumn_ = 1; worksheet.Cells[tableColumn_][1].value2 != null; tableColumn_++)
                            { }
                        }

                    }
                    finally
                    {
                        workbook.Close();
                        Marshal.ReleaseComObject(xlApp);
                        Marshal.ReleaseComObject(workbook);
                        Marshal.ReleaseComObject(worksheet);
                        Marshal.ReleaseComObject(range);
                        xlApp = null;
                        workbook = null;
                        worksheet = null;
                        range = null;
                    }
                }
            }
            finally
            {
                tiaPortal.Dispose();
            }
        }

        static void CreateScreenItem(HmiScreen screen, Dictionary<string, object> propertyNameValues)
        {
            string sName = propertyNameValues["Name"].ToString();
            propertyNameValues.Remove("Name");
            string sType = propertyNameValues["Type"].ToString();
            propertyNameValues.Remove("Type");
            Console.WriteLine("CreateScreenItem: " + sName + " of type " + sType);
            Type type = null;
            if (sType == "HmiLine" || sType == "HmiPolyline" || sType == "HmiPolygon" || sType == "HmiEllipse" || sType == "HmiEllipseSegment"
                || sType == "HmiCircleSegment" || sType == "HmiEllipticalArc" || sType == "HmiCircularArc" || sType == "HmiCircle" || sType == "HmiRectangle"
                || sType == "HmiGraphicView")
            {
                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Shapes." + sType + ", Siemens.Engineering");
            }
            else if (sType == "HmiIOField" || sType == "HmiButton" || sType == "HmiToggleSwitch" || sType == "HmiCheckBoxGroup" || sType == "HmiBar"
                || sType == "HmiGauge" || sType == "HmiSlider" || sType == "HmiRadioButtonGroup" || sType == "HmiListBox" || sType == "HmiClock"
                || sType == "HmiTextBox")
            {
                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Widgets." + sType + ", Siemens.Engineering");
            }
            else if (sType == "HmiAlarmControl" || sType == "HmiMediaControl" || sType == "HmiTrendControl" || sType == "HmiTrendCompanion"
                || sType == "HmiProcessControl" || sType == "HmiFunctionTrendControl" || sType == "HmiWebControl" || sType == "HmiDetailedParameterControl" || sType == "HmiFaceplateContainer")
            {
                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Controls." + sType + ", Siemens.Engineering");
            }
            else if (sType == "HmiScreenWindow")
            {
                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Screens." + sType + ", Siemens.Engineering");
            }
            else if (sType == "HmiScreen")
            {
                // to prevent returning from this function without doing nothing
            }
            else if (sType.ToLower() == "pause")
            {
                Console.WriteLine("Progress paused due to command '" + sType + "'. Please hit any key to continue...");
                Console.Read();
                return;
            }
            else {
                Console.WriteLine("ScreenItem with type " + sType + " is not implemented yet!");
                return;
            }

            UIBase screenItem = screen;
            if (sType != "HmiScreen")
            {
                screenItem = screen.ScreenItems.Find(sName);
            }
            if (screenItem == null)
            {
                MethodInfo createMethod = typeof(HmiScreenItemBaseComposition).GetMethod("Create", BindingFlags.Public | BindingFlags.Instance, null, CallingConventions.Any, new Type[] { typeof(string) }, null);
                MethodInfo generic = createMethod.MakeGenericMethod(type);
                screenItem = (HmiScreenItemBase)generic.Invoke(screen.ScreenItems, new object[] { sName });
            }

            foreach (var propertyNameValue in propertyNameValues)
            {
                if (propertyNameValue.Value.ToString() == "")
                {
                    continue;
                }
                else { 
                    Console.WriteLine("Will try to set Property '" + propertyNameValue.Key + "' with value '" + propertyNameValue.Value + "'.");
                    try
                    {
                        //cover # Attributes
                        if (propertyNameValue.Key.Contains("Property") && propertyNameValue.Key.Contains("Events"))
                        {
                            SetChangeEvent(screenItem, propertyNameValue.Key, propertyNameValue.Value);
                        }
                        else if (propertyNameValue.Key.Contains("Events"))
                        {
                            string key = propertyNameValue.Key.Split('.')[1] + '.' + propertyNameValue.Key.Split('.')[2];
                            SetEvent(screenItem, key, propertyNameValue.Value);
                        }
                        else if (propertyNameValue.Key.Contains("Dynamization"))
                        {
                            SetDynamization(screenItem, propertyNameValue.Key, propertyNameValue.Value);
                        }
                        else
                        {
                            SetPropertyRecursive(propertyNameValue.Key, propertyNameValue.Value.ToString(), screenItem);
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Cannot set property '" + propertyNameValue.Key + "' with value '" + propertyNameValue.Value + "'.");
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="screenItem"></param>
        /// <param name="key">Down#ScriptCode or Down#Async</param>
        /// <param name="value"></param>
        private static void SetEvent(UIBase screenItem, string key, object value)
        {
            var comp = (screenItem as IEngineeringObject).GetComposition("EventHandlers") as IEngineeringComposition;
            MethodBase createMethod = comp.GetType().GetMethod("Create");
            MethodBase findMethod = comp.GetType().GetMethod("Find");

            //split key
            string[] keys = key.Split('.');
            foreach (var item in Enum.GetValues(createMethod.GetParameters()[0].ParameterType))
            {
                if (item.ToString() == keys[0])
                {
                    var eventHandler = findMethod.Invoke(comp, new object[] { item }) ?? createMethod.Invoke(comp, new object[] { item });
                    IEngineeringObject event_ = (eventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                    SetPropertyRecursive(keys[1], value.ToString(), event_);
                }
            }
        }

        private static void SetChangeEvent(UIBase screenItem, string key, object value)
        {
            

            var comp = (screenItem as IEngineeringObject).GetComposition("PropertyEventHandlers") as IEngineeringComposition;
            MethodBase createMethod = comp.GetType().GetMethod("Create");
            MethodBase findMethod = comp.GetType().GetMethod("Find");
            string key1 = key.Split('.')[0];
            string key2 = key.Split('.')[2];

            foreach (var item in Enum.GetValues(createMethod.GetParameters()[1].ParameterType))
            {
                if (key.Split('.')[1] == "PropertyQualityCodeEvents")
                {
                        IEngineeringObject propEventHandler = null;
                        propEventHandler = screenItem.PropertyEventHandlers.Create("ProcessValue", Siemens.Engineering.HmiUnified.UI.Events.PropertyEventType.QualityCodeChange);
                        IEngineeringObject propEvent = (propEventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                        SetPropertyRecursive(key2, value.ToString(), propEvent);

                    //else
                    //{
                    //    var propEventHandler = findMethod.Invoke(comp, new object[] { key1, item }) ?? createMethod.Invoke(comp, new object[] { key1, item });
                    //    IEngineeringObject propEvent = (propEventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                    //    SetPropertyRecursive(key2, value.ToString(), propEvent);
                    //}

                }
                else if (item.ToString() == "Change")
                {
                        var propEventHandler = findMethod.Invoke(comp, new object[] { key1, item }) ?? createMethod.Invoke(comp, new object[] { key1, item });
                        IEngineeringObject propEvent = (propEventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                        SetPropertyRecursive(key2, value.ToString(), propEvent);
                }

            }
        }

        //string oldPropertyName = "";
        private static void SetDynamization(UIBase screenItem, string key, object value)
        {
            var comp = (screenItem as IEngineeringObject).GetComposition("Dynamizations") as IEngineeringComposition;
            MethodBase findMethod = comp.GetType().GetMethod("Find");

            var keyLength = key.Split('.').Length;
            string[] keys = new string[keyLength];
            int dynStart = 0;
            for (int i = 0; i < keyLength; i++)
            {
                keys[i] = key.Split('.')[i];
                //find where the dynamization starts
                if (keys[i].Contains("Dynamization"))
                {
                    dynStart = i;
                }
            }
            

            var findDyn = findMethod.Invoke(comp, new object[] { keys[0] });
            
            // todo: make generic
            if (keys[dynStart].StartsWith("Script") && findMethod != null)
            {
                if (Globals.oldPropertyName != keys[dynStart - 1])
                {
                    Globals.scriptTemp = null;
                }

                if (screenItem is Siemens.Engineering.HmiUnified.UI.Controls.HmiFaceplateContainer)
                {
                    var fpContainer = screenItem as HmiFaceplateContainer;
                    var str = keys[0].ToString().Split('[')[1];
                    str = str.Split(']')[0];
                    int number = Convert.ToInt32(str);
                    if (Globals.oldPropertyName != keys[dynStart - 1])
                    {
                        Globals.scriptTemp = fpContainer.Interface[number].Dynamizations.Create<ScriptDynamization>(keys[dynStart - 1]);
                    }
                }
                
                else
                {
                    Globals.scriptTemp = (ScriptDynamization)findMethod.Invoke(comp, new object[] { keys[dynStart - 1] }) ?? screenItem.Dynamizations.Create<ScriptDynamization>(keys[dynStart - 1]);
                }
                if (keys[keyLength - 1].Contains("GlobalDefinition"))
                {
                    Globals.scriptTemp.GlobalDefinitionAreaScriptCode = value.ToString();
                }
                else if (keys[keyLength - 1].Contains("ScriptCode"))
                {
                    Globals.scriptTemp.ScriptCode = value.ToString();
                }

                else if (keys[keyLength - 1].Contains("Type"))
                {
                    Globals.scriptTemp.Trigger.Type = TriggerType.Tags;
                }

                else if (keys[keyLength - 1].Contains("Tags"))
                {
                    var tagCount = value.ToString().Split(',').Length;
                    List<string> tags = new List<string>();
                    for (int i = 0; i < tagCount; i++)
                    {
                        tags.Add(value.ToString().Split(',')[i]);
                    }
                    Globals.scriptTemp.Trigger.Tags = tags;
                }
                else
                {
                    SetPropertyRecursive(keys[dynStart + 1], value.ToString(), (Globals.scriptTemp as IEngineeringObject));
                }
            }
            else if (keys[dynStart].StartsWith("Tag") && findMethod != null)
            {
                TagDynamization temp = null;
                if (screenItem is Siemens.Engineering.HmiUnified.UI.Controls.HmiFaceplateContainer)
                {
                    var fpContainer = screenItem as HmiFaceplateContainer;
                    var str = keys[0].ToString().Split('[')[1];
                    str = str.Split(']')[0];
                    int number = Convert.ToInt32(str);
                    temp = fpContainer.Interface[number].Dynamizations.Create<TagDynamization>(keys[dynStart - 1]);
                }
                else
                {
                    temp = (TagDynamization)findMethod.Invoke(comp, new object[] { keys[dynStart - 1] }) ?? screenItem.Dynamizations.Create<TagDynamization>(keys[dynStart - 1]);
                }
                SetPropertyRecursive(keys[dynStart + 1], value.ToString(), (temp as IEngineeringObject));
            }
            else if (keys[dynStart].StartsWith("ResourceList") && findMethod != null)
            {
                ResourceListDynamization temp = null;
                if (screenItem is Siemens.Engineering.HmiUnified.UI.Controls.HmiFaceplateContainer)
                {
                    var fpContainer = screenItem as HmiFaceplateContainer;
                    var str = keys[0].ToString().Split('[')[1];
                    str = str.Split(']')[0];
                    int number = Convert.ToInt32(str);
                    temp = fpContainer.Interface[number].Dynamizations.Create<ResourceListDynamization>(keys[dynStart - 1]);
                }
                else
                {
                    temp = (ResourceListDynamization)findMethod.Invoke(comp, new object[] { keys[dynStart - 1] }) ?? screenItem.Dynamizations.Create<ResourceListDynamization>(keys[dynStart - 1]);
                }
                    SetPropertyRecursive(keys[dynStart + 1], value.ToString(), (temp as IEngineeringObject));
            }
            else if (keys[dynStart].StartsWith("Flashing") && findMethod != null)
            {
                FlashingDynamization temp = null;
                if (screenItem is Siemens.Engineering.HmiUnified.UI.Controls.HmiFaceplateContainer)
                {
                    var fpContainer = screenItem as HmiFaceplateContainer;
                    var str = keys[0].ToString().Split('[')[1];
                    str = str.Split(']')[0];
                    int number = Convert.ToInt32(str);
                    temp = fpContainer.Interface[number].Dynamizations.Create<FlashingDynamization>(keys[dynStart - 1]);
                }
                else
                {
                    temp = (FlashingDynamization)findMethod.Invoke(comp, new object[] { keys[dynStart - 1] }) ?? screenItem.Dynamizations.Create<FlashingDynamization>(keys[dynStart - 1]);
                }
                    SetPropertyRecursive(keys[dynStart + 1], value.ToString(), (temp as IEngineeringObject));
            }
            Globals.oldPropertyName = keys[dynStart - 1];
        }

        static public void SetMyAttributesSimpleTypes(string keyToSet, object valueToSet, IEngineeringObject obj)
        {
            object _attr;
            // To reach the MultilingualText handling branch later, we need to get past the
            // EngineeringNotSupportedException that trying to access the attribute raises.
            try
            {
                _attr = obj.GetAttribute(keyToSet);
            }
            catch (Siemens.Engineering.EngineeringNotSupportedException ex)
            {
                Console.WriteLine("Cannot access {0} using GetAttribute(): {1}", keyToSet, ex.Message);
                _attr = null;
            }
            Type _type = null;
            if (_attr != null)
            {
                _type = _attr.GetType();
            }
            else
            {
                var attrInfos = obj.GetAttributeInfos();
                var attrInfo = attrInfos.Where(x => x.Name == keyToSet);
                if (attrInfo.Any())
                {
                    _type = attrInfo.First().SupportedTypes.First();
                }
            }

            object attrVal = null;
            if (_type != null && _type.BaseType == typeof(Enum))
            {
                attrVal = Enum.Parse(_type, valueToSet.ToString());
            }
            else if (_type != null && _type.Name == "Color")
            {
                var hexColor = new ColorConverter();
                attrVal = (Color)hexColor.ConvertFromString(valueToSet.ToString().ToUpper());
            }
            else if (keyToSet == "InitialAddress")
            {
                attrVal = valueToSet.ToString().Substring(0, valueToSet.ToString().Length - 1);
            }
            else if (_type != null && _type.Name == "MultilingualText")
            {
                var multiLingText = obj.GetAttribute(keyToSet) as MultilingualText;
                obj = multiLingText.Items.FirstOrDefault();

                if (obj == null)
                {
                    Console.WriteLine("Cannot find a language for the text property '" + keyToSet + "'.");
                    return;
                }
                keyToSet = "Text";
                attrVal = valueToSet.ToString();
            }
            else if (obj.GetType().Name == "MultilingualText")
            {
                var multiLingText = obj as MultilingualText;
                obj = multiLingText.Items.First(x => x.Language.Culture.Name == keyToSet);
                if (obj == null)
                {
                    Console.WriteLine("Cannot find a language for the text property '" + keyToSet + "'.");
                    return;
                }
                keyToSet = "Text";
                attrVal = valueToSet.ToString();
            }
            else if (keyToSet == "Tags")
            {
                attrVal = (valueToSet as List<object>).Select(i => i.ToString()).ToList();
            }
            else
            {
                if (_type != null) attrVal = Convert.ChangeType(valueToSet, _type);
            }

            try
            {
                obj.SetAttribute(keyToSet.ToString(), attrVal);
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }

        public static void SetPropertyRecursive(string key, string value, IEngineeringObject relevantTag)
        {
            if (key.Contains("."))
            {
                IEngineeringObject deeperObj = null;
                List<string> keySplit = key.Split('.').ToList();
                if (keySplit[0].EndsWith("]")) //trerndAreas[0]
                {
                    var compositionsName = keySplit[0].Split('[')[0];
                    var composition = relevantTag.GetComposition(compositionsName) as IEngineeringComposition;
                    var indexString = keySplit[0].Split('[')[1];
                    int index = int.Parse(indexString.Split(']')[0]);
                    int count = composition.Count;
                    while (count <= index)
                    {
                        deeperObj = (composition is HmiPointComposition) ? (composition as HmiPointComposition).Create(0, 0) : composition.Create(composition[0].GetType(), null);
                        count++;
                    }
                    if (deeperObj == null)
                    {
                        deeperObj = composition[index];
                    }
                }
                else
                {
                    deeperObj = relevantTag.GetAttribute(keySplit[0]) as IEngineeringObject;
                }
                
                keySplit.RemoveAt(0);
                string deeperKey = string.Join(".", keySplit);
                SetPropertyRecursive(deeperKey, value, deeperObj);
            }
            else
            {
                SetMyAttributesSimpleTypes(key, value, relevantTag);
            }
        }
        static IEnumerable<HmiSoftware> GetHmiSoftwares(ref TiaPortal tiaPortal)
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
            return
                from device in tiaPortalProject.Devices
                from deviceItem in device.DeviceItems
                let softwareContainer = deviceItem.GetService<SoftwareContainer>()
                where softwareContainer?.Software is HmiSoftware
                select softwareContainer.Software as HmiSoftware;
        }

    }
}
