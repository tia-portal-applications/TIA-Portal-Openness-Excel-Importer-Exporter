using System;
using System.Collections.Generic;
using System.Linq;
using Siemens.Engineering.HmiUnified;
using System.Drawing;
using Siemens.Engineering;
using Siemens.Engineering.HmiUnified.UI.Dynamization;
using Siemens.Engineering.HmiUnified.UI;
using Siemens.Engineering.HmiUnified.UI.Screens;
using System.Reflection;
using Siemens.Engineering.HmiUnified.UI.Base;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Script;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Flashing;
using Siemens.Engineering.HmiUnified.UI.Parts;
using Siemens.Engineering.HmiUnified.UI.ScreenGroup;

namespace ExcelExporter
{
    class Screen : ITiaObject
    {
        private static IEngineeringObject dyn;
        private string screenName = "-all";
        private List<string> definedAttributes = null;
        public Screen(string screenName = "-all", List<string> definedAttributes = null)
        {
            this.screenName = screenName;
            this.definedAttributes = definedAttributes;
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

        public Dictionary<string, List<Dictionary<string, object>>> Export(HmiSoftware hmiSoftware)
        {
            var list = new Dictionary<string, List<Dictionary<string, object>>>();
            IEnumerable<HmiScreen> allScreens = GetScreens(hmiSoftware);
            if (screenName != "-all")
            {
                allScreens = new List<HmiScreen>() { allScreens.First(s => s.Name == screenName) };
            }
            //get main Screen
            //IEnumerable<HmiScreen> mainScreen = ;
            foreach (var screen in allScreens)
            {
                
                var listScreenItems = new List<Dictionary<string, object>>();
                listScreenItems.Add(Helper.GetAllMyAttributes(screen, ParseScreen, definedAttributes));

                foreach (var screenItem in screen.ScreenItems)
                {
                    var screenItem_ = Helper.GetAllMyAttributes(screenItem, ParseScreen, definedAttributes);
                    listScreenItems.Add(screenItem_);
                }
                list.Add(screen.Name, listScreenItems);
            }
            return list;
        }

        /// <summary>
        /// Change the hierarchy of Dynamization, Events and PropertyEvents in Screens and add the type and MultiLingualText attributes in the YAML-File
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private bool ParseScreen(ParseHelperClass helperclass)
        {
            var obj = helperclass.obj;
            var dict = helperclass.dictionairy;
            if(obj is HmiScreenBase || obj is HmiScreenItemBase)
            {
                dict.Add("Type", obj.GetType().Name);
            }

            if (obj is UIBase)
            {
                GetDynamizationsAndEvents(obj, ref dict);
            }

            if (obj is MultilingualText)
            {
                Helper.GetMultilingualTextItems(obj, ref dict);    
            }            
            return true;
        }


        private void GetDynamizationsAndEvents(IEngineeringObject obj, ref Dictionary<string, object> dict)
        {
            var dynamizations = (obj as UIBase).Dynamizations;
            if (dynamizations.Count != 0 && (this.definedAttributes == null || this.definedAttributes.Find(definedAttr => definedAttr.Split('.').Where(x => x == "Dynamizations").Count() > 0) != null))
            {
                var dyns = new Dictionary<string, object>();
                foreach (var dyn in (obj as UIBase).Dynamizations)
                {
                    var dynType = "";
                    var attrKeys = from attributeInfo in dyn.GetType().GetProperties()
                                   where attributeInfo.CanWrite
                                   select attributeInfo.Name;
                    var attrProps = dyn.GetAttributes(attrKeys);
                    var attrDict = attrKeys.Zip(attrProps, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);

                    if (dyn.DynamizationType == DynamizationType.Script)
                    {
                        dynType = "ScriptDynamization";
                        var trigger = dyn.GetType().GetProperty("Trigger").GetValue(dyn);
                        var triggerDict = new Dictionary<string, object>();

                        triggerDict.Add("CustomDuration", trigger.GetType().GetProperty("CustomDuration").GetValue(trigger));
                        triggerDict.Add("Type", trigger.GetType().GetProperty("Type").GetValue(trigger));
                        triggerDict.Add("Tags", trigger.GetType().GetProperty("Tags").GetValue(trigger));

                        attrDict.Add("Trigger", triggerDict as object);
                    }
                    else if (dyn.DynamizationType == DynamizationType.Flashing)
                    {
                        dynType = "FlashingDynamization";
                        var tempDict = new Dictionary<string, object>();
                        foreach (var item in attrDict)
                        {
                            if (item.Value.GetType().Name == "Color")
                            {
                                var color = (Color)item.Value;
                                tempDict.Add(item.Key, "0x" + color.A.ToString("X2") + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2"));
                            }
                            else
                            {
                                tempDict.Add(item.Key, item.Value);
                            }
                        }
                        attrDict.Clear();
                        attrDict = attrDict.Concat(tempDict).ToDictionary(x => x.Key, x => x.Value);
                    }
                    else if(dyn.DynamizationType == DynamizationType.Tag)
                    {
                        dynType = "TagDynamization";
                    }
                    else if (dyn.DynamizationType == DynamizationType.ResourceList)
                    {
                        dynType = "ResourceListDynamization";
                    }

                    var dictPropertyName = dyn.PropertyName +"."+ dynType;

                    dict.Add(dictPropertyName, attrDict as Object);
                }
                
            }

            if (this.definedAttributes == null || this.definedAttributes.Find(definedAttr => definedAttr.Split('.').Where(x => x == "Events").Count() > 0) != null)
            { 
                try
                {
                    var event_ = obj.GetComposition("EventHandlers");
                    var events = new Dictionary<string, object>();
                    foreach (IEngineeringObject eve in obj.GetComposition("EventHandlers") as IEngineeringComposition)
                    {
                        IEngineeringObject script = eve.GetAttribute("Script") as IEngineeringObject;

                        var attrKeys = from attributeInfo in script.GetType().GetProperties()
                                       where attributeInfo.CanWrite
                                       select attributeInfo.Name;
                        var attrProps = script.GetAttributes(attrKeys);
                        var attrDict = attrKeys.Zip(attrProps, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);

                        string eventType = eve.GetAttribute("EventType").ToString();

                        events.Add(eventType, attrDict);
                    }

                    if (events.Count != 0)
                    {
                        dict.Add("Events", events as Object);
                    }
                }
                catch (Exception ex) { /*Console.WriteLine("In some obj is no EventHandler available: " + ex.Message);*/ }
            }

            if (!(obj is HmiFaceplateInterface) && !(obj is HmiCustomControlInterface)) // Faceplate and CWC interfaces do not have property event handlers or quality event handlers. TIA Portal crashes when you access it via TIA Openness.
            {
                var propertyEventHandler = (obj as UIBase).PropertyEventHandlers.Where(x => x.EventType == Siemens.Engineering.HmiUnified.UI.Events.PropertyEventType.Change);
                if (propertyEventHandler.Count() != 0 && (this.definedAttributes == null || this.definedAttributes.Find(definedAttr => definedAttr.Split('.').Where(x => x == "PropertyEvents").Count() > 0) != null))
                {
                    var propEves = new Dictionary<string, object>();
                    foreach (var propEve in propertyEventHandler)
                    {
                        IEngineeringObject script = propEve.Script as IEngineeringObject;

                        var attrKeys = from attributeInfo in script.GetType().GetProperties()
                                       where attributeInfo.CanWrite
                                       select attributeInfo.Name;
                        var attrProps = script.GetAttributes(attrKeys);
                        var attrDict = attrKeys.Zip(attrProps, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);

                        string propEventType = propEve.PropertyName.ToString();

                        propEves.Add(propEventType, attrDict);
                    }
                    dict.Add("PropertyEvents", propEves as Object);
                }

                var propertyQualityEventHandler = (obj as UIBase).PropertyEventHandlers.Where(x => x.EventType == Siemens.Engineering.HmiUnified.UI.Events.PropertyEventType.QualityCodeChange);
                if (propertyQualityEventHandler.Count() != 0 && (this.definedAttributes == null || this.definedAttributes.Find(definedAttr => definedAttr.Split('.')[0] == "PropertyQualityCodeEvents") != null))
                {
                    var propEves = new Dictionary<string, object>();
                    foreach (var propEve in propertyQualityEventHandler)
                    {
                        IEngineeringObject script = propEve.Script as IEngineeringObject;

                        var attrKeys = from attributeInfo in script.GetType().GetProperties()
                                       where attributeInfo.CanWrite
                                       select attributeInfo.Name;
                        var attrProps = script.GetAttributes(attrKeys);
                        var attrDict = attrKeys.Zip(attrProps, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);

                        string propEventType = propEve.PropertyName.ToString();

                        propEves.Add(propEventType, attrDict);
                        dict.Add(propEventType+"."+"PropertyQualityCodeEvents", attrDict);
                    }
                    
                }
            }
        }

        public void Import(HmiSoftware hmiSoftware, List<object> data)
        {
            IEnumerable<HmiScreen> allScreens = GetScreens(hmiSoftware);
            HmiScreen _screen = null;
            foreach (var topLevel in data)
            {
                foreach (var middleLevel in topLevel as Dictionary<object, object>)
                {
                    foreach (var value in middleLevel.Value as List<object>)
                    {
                        var dataTree = (value as Dictionary<object, object>).ToDictionary(k => k.Key.ToString(), k => k.Value);

                        if (dataTree["Type"].ToString() == "HmiScreen")
                        {
                            _screen = allScreens.FirstOrDefault(s => s.Name == dataTree["Name"].ToString()) ?? hmiSoftware.Screens.Create(dataTree["Name"].ToString());
                            dataTree.Remove("Type");
                            dataTree.Remove("Name");

                            Helper.SetAllMyAttributes(dataTree, _screen, ReParseScreen);
                        }
                        else 
                        {
                            Type type = null;
                            if (dataTree["Type"].ToString() == "HmiLine" || dataTree["Type"].ToString() == "HmiPolyline" || dataTree["Type"].ToString() == "HmiPolygon" || dataTree["Type"].ToString() == "HmiEllipse" || dataTree["Type"].ToString() == "HmiEllipseSegment"
                                || dataTree["Type"].ToString() == "HmiCircleSegment" || dataTree["Type"].ToString() == "HmiEllipticalArc" || dataTree["Type"].ToString() == "HmiCircularArc" || dataTree["Type"].ToString() == "HmiCircle" || dataTree["Type"].ToString() == "HmiRectangle"
                                || dataTree["Type"].ToString() == "HmiGraphicView") 
                            {
                                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Shapes." + dataTree["Type"].ToString() + ", Siemens.Engineering");
                            }
                            else if (dataTree["Type"].ToString() == "HmiIOField" || dataTree["Type"].ToString() == "HmiButton" || dataTree["Type"].ToString() == "HmiToggleSwitch" || dataTree["Type"].ToString() == "HmiCheckBoxGroup" || dataTree["Type"].ToString() == "HmiBar"
                                || dataTree["Type"].ToString() == "HmiGauge" || dataTree["Type"].ToString() == "HmiSlider" || dataTree["Type"].ToString() == "HmiRadioButtonGroup" || dataTree["Type"].ToString() == "HmiListBox" || dataTree["Type"].ToString() == "HmiClock"
                                || dataTree["Type"].ToString() == "HmiTextBox")
                            {
                                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Widgets." + dataTree["Type"].ToString() + ", Siemens.Engineering");
                            }
                            else if (dataTree["Type"].ToString() == "HmiAlarmControl" || dataTree["Type"].ToString() == "HmiMediaControl" || dataTree["Type"].ToString() == "HmiTrendControl" || dataTree["Type"].ToString() == "HmiTrendCompanion"
                                || dataTree["Type"].ToString() == "HmiProcessControl" || dataTree["Type"].ToString() == "HmiFunctionTrendControl" || dataTree["Type"].ToString() == "HmiWebControl" || dataTree["Type"].ToString() == "HmiDetailedParameterControl" || dataTree["Type"].ToString() == "HmiFaceplateContainer")
                            {
                                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Controls." + dataTree["Type"].ToString() + ", Siemens.Engineering");
                            }
                            else if (dataTree["Type"].ToString() == "HmiScreenWindow")
                            {
                                type = Type.GetType("Siemens.Engineering.HmiUnified.UI.Screens." + dataTree["Type"].ToString() + ", Siemens.Engineering");
                            }
                            else { Console.WriteLine("ScreenItem with type " + dataTree["Type"].ToString() + " is not implemented yet!"); }

                            var screenItem = _screen.ScreenItems.Find(dataTree["Name"].ToString());
                            if (screenItem == null) //not availabe in screen
                            {
                                MethodInfo createMethod = typeof(HmiScreenItemBaseComposition).GetMethod("Create", BindingFlags.Public | BindingFlags.Instance, null, CallingConventions.Any, new Type[] { typeof(string) }, null);
                                MethodInfo generic = createMethod.MakeGenericMethod(type);

                                var newItem = (IEngineeringObject)generic.Invoke(_screen.ScreenItems, new object[] { dataTree["Name"].ToString() });

                                dataTree.Remove("Type");
                                dataTree.Remove("Name");
                                Helper.SetAllMyAttributes(dataTree, newItem, ReParseScreenItem);
                            }
                            else
                            {
                                dataTree.Remove("Type");
                                dataTree.Remove("Name");
                                Helper.SetAllMyAttributes(dataTree, screenItem as IEngineeringObject, ReParseScreenItem);
                            }
                        }
                    }
                }
            }
        }
        private IEngineeringObject ReParseScreen(IEngineeringObject obj, string keyDict, object valueDict)
        {
            if (keyDict == "Dynamizations")
            {
                foreach (var attr in (valueDict as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value))
                {
                    var dyns = (attr.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value);

                    var comp = obj.GetComposition("Dynamizations") as IEngineeringComposition;

                    // MethodBase createMethod = comp.GetType().GetMethod("Create"); //Create[T] => [TagDynamization]???
                    //var ets = createMethod.Invoke(comp, new object[] { attr.Key });

                    MethodBase findMethod = comp.GetType().GetMethod("Find");
                    var findDyn = findMethod.Invoke(comp, new object[] { attr.Key });


                    if (dyns.ContainsKey("ScriptCode") && findMethod != null)
                    {
                        var scriptDyn = (obj as HmiScreen).Dynamizations.Create<ScriptDynamization>(attr.Key);
                        dyn = (scriptDyn as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreen);
                    }
                    else if (dyns.ContainsKey("ResourceList") && findMethod != null)
                    {
                        var resourcelistDyn = (obj as HmiScreen).Dynamizations.Create<ResourceListDynamization>(attr.Key);
                        dyn = (resourcelistDyn as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreen);
                    }
                    else if (dyns.ContainsKey("FlashingCondition") && findMethod != null)
                    {
                        var flashingDyn = (obj as HmiScreen).Dynamizations.Create<FlashingDynamization>(attr.Key);
                        dyn = (flashingDyn as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreen);
                    }
                    else if (dyns.ContainsKey("UseIndirectAddressing") && findMethod != null)
                    {
                        var tagDyn = (obj as HmiScreen).Dynamizations.Create<TagDynamization>(attr.Key);
                        dyn = (tagDyn as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreen);
                    }
                }
            }
            else if (keyDict == "Events")
            {
                foreach (var attr in (valueDict as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value))
                {
                    var events = (attr.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value);

                    var comp = obj.GetComposition("EventHandlers") as IEngineeringComposition;
                    MethodBase createMethod = comp.GetType().GetMethod("Create");                  
                    MethodBase findMethod = comp.GetType().GetMethod("Find");

                    foreach (var item in Enum.GetValues(createMethod.GetParameters()[0].ParameterType))
                    {
                        if(item.ToString() == attr.Key)
                        {
                            var eventHandler = findMethod.Invoke(comp, new object[] { item }) ?? createMethod.Invoke(comp, new object[] { item });
                            IEngineeringObject event_ = (eventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                            Helper.SetAllMyAttributes(events, event_, ReParseScreen);
                        }
                    }
                }
            }
            else if (keyDict == "PropertyEvents")
            {
                foreach (var attr in (valueDict as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value))
                {
                    var propEvents = (attr.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value);

                    var comp = obj.GetComposition("PropertyEventHandlers") as IEngineeringComposition;
                    MethodBase createMethod = comp.GetType().GetMethod("Create");
                    MethodBase findMethod = comp.GetType().GetMethod("Find");

                    foreach (var item in Enum.GetValues(createMethod.GetParameters()[1].ParameterType))
                    {
                        if (item.ToString() == "Change")
                        {
                            var propEventHandler = findMethod.Invoke(comp, new object[] { attr.Key, item }) ?? createMethod.Invoke(comp, new object[] { attr.Key, item });
                            IEngineeringObject propEvent = (propEventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                            Helper.SetAllMyAttributes(propEvents, propEvent, ReParseScreen);
                        }
                    }
                }
            }                                
            return null;
        }
        private IEngineeringObject ReParseScreenItem(IEngineeringObject obj, string keyDict, object valueDict)
        {
            if (keyDict == "Dynamizations")
            {
                foreach (var attr in (valueDict as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value))
                {
                    var dyns = (attr.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value);

                    var comp = obj.GetComposition("Dynamizations") as IEngineeringComposition;

                    MethodBase findMethod = comp.GetType().GetMethod("Find");
                    var findDyn = findMethod.Invoke(comp, new object[] { attr.Key });


                    if (dyns.ContainsKey("ScriptCode") && findMethod != null)
                    {
                        ScriptDynamization temp = null;
                        if(obj is HmiScreenPartBase)
                        {
                            temp = (obj as HmiScreenPartBase).Dynamizations.Create<ScriptDynamization>(attr.Key); 
                        }
                        else
                        {
                            temp = (obj as HmiScreenItemBase).Dynamizations.Create<ScriptDynamization>(attr.Key);
                        }
                        dyn = (temp as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreenItem);
                    }
                    else if (dyns.ContainsKey("UseIndirectAddressing") && findMethod != null)
                    {
                        TagDynamization temp = null;
                        if (obj is HmiScreenPartBase) 
                        {
                            temp = (obj as HmiScreenPartBase).Dynamizations.Create<TagDynamization>(attr.Key);
                        }
                        else
                        {
                            temp = (obj as HmiScreenItemBase).Dynamizations.Create<TagDynamization>(attr.Key);
                        }
                        dyn = (temp as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreenItem);
                    }
                    else if (dyns.ContainsKey("ResourceList") && findMethod != null)
                    {
                        ResourceListDynamization temp = null;
                        if (obj is HmiScreenPartBase)
                        {
                            temp = (obj as HmiScreenPartBase).Dynamizations.Create<ResourceListDynamization>(attr.Key);
                        }
                        else
                        {
                            temp = (obj as HmiScreenItemBase).Dynamizations.Create<ResourceListDynamization>(attr.Key);
                        }
                        dyn = (temp as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreenItem);
                    }
                    else if (dyns.ContainsKey("FlashingCondition") && findMethod != null)
                    {
                        FlashingDynamization temp = null;
                        if (obj is HmiScreenPartBase)
                        {
                            temp = (obj as HmiScreenPartBase).Dynamizations.Create<FlashingDynamization>(attr.Key);
                        }
                        else
                        {
                            temp = (obj as HmiScreenItemBase).Dynamizations.Create<FlashingDynamization>(attr.Key);
                        }
                        dyn = (temp as IEngineeringObject);
                        Helper.SetAllMyAttributes(dyns, dyn, ReParseScreenItem);
                    }
                }
            }
            else if(keyDict == "Events")
            {
                foreach (var attr in (valueDict as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value))
                {
                    var events = (attr.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value);

                    var comp = obj.GetComposition("EventHandlers") as IEngineeringComposition;
                    MethodBase createMethod = comp.GetType().GetMethod("Create");
                    MethodBase findMethod = comp.GetType().GetMethod("Find");

                    foreach (var item in Enum.GetValues(createMethod.GetParameters()[0].ParameterType))
                    {
                        if (item.ToString() == attr.Key)
                        {
                            var eventHandler = findMethod.Invoke(comp, new object[] { item }) ?? createMethod.Invoke(comp, new object[] { item });
                            IEngineeringObject event_ = (eventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                            Helper.SetAllMyAttributes(events, event_, ReParseScreenItem);
                        }
                    }
                }
            }
            else if(keyDict == "PropertyEvents")
            {
                foreach (var attr in (valueDict as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value))
                {
                    var propEvents = (attr.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value);

                    var comp = obj.GetComposition("PropertyEventHandlers") as IEngineeringComposition;
                    MethodBase createMethod = comp.GetType().GetMethod("Create");
                    MethodBase findMethod = comp.GetType().GetMethod("Find");

                    foreach (var item in Enum.GetValues(createMethod.GetParameters()[1].ParameterType))
                    {
                        if (item.ToString() == "Change")
                        {
                            var propEventHandler = findMethod.Invoke(comp, new object[] { attr.Key, item }) ?? createMethod.Invoke(comp, new object[] { attr.Key, item });
                            IEngineeringObject propEvent = (propEventHandler as IEngineeringObject).GetAttribute("Script") as IEngineeringObject;
                            Helper.SetAllMyAttributes(propEvents, propEvent, ReParseScreenItem);
                        }
                    }
                }
            }
                return null;
        }
        static public Dictionary<string, object> SetTrigger(Dictionary<string, object> dataTree, IEngineeringObject obj)
        {
            foreach (var _trigger in dataTree)
            {
                if (_trigger.Key.ToString() == "Trigger")
                {
                    try
                    {
                        var node = obj.GetAttribute(_trigger.Key) as IEngineeringObject;
                        Helper.SetAllMyAttributes((_trigger.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value), node, null);
                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); }
                }
            }
            return dataTree;
        }
    }
}
