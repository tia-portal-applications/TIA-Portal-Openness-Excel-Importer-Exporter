using Siemens.Engineering;
using Siemens.Engineering.HmiUnified.HmiAlarm;
using Siemens.Engineering.HmiUnified.HmiTags;
using Siemens.Engineering.HmiUnified.UI;
using Siemens.Engineering.HmiUnified.UI.Dynamization;
using Siemens.Engineering.HmiUnified.UI.Dynamization.Script;
using Siemens.Engineering.HmiUnified.UI.Screens;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace ExcelExporter
{
    class ParseHelperClass
    {
        public ParseHelperClass(IEngineeringObject obj, Dictionary<string, object> dictionairy)
        {
            this.obj = obj;
            this.dictionairy = dictionairy;
        }
        public IEngineeringObject obj;
        public Dictionary<string, object> dictionairy;
    }
    class Helper
    {
        static public Dictionary<string, object> GetAllMyAttributes(IEngineeringObject obj, Func<ParseHelperClass, bool> parse, List<string> definedAttributes = null, string fullName = "")
        {
            //Leaves
                var attrKeys = from attributeInfo in obj.GetAttributeInfos()
                           where Helper.IsSimple(attributeInfo.SupportedTypes.FirstOrDefault()) && attributeInfo.AccessMode.ToString() == "ReadWrite" && (definedAttributes == null ||fullName + attributeInfo.Name == "Name" || definedAttributes.Find(x => x == fullName + attributeInfo.Name) != null)
                           select attributeInfo.Name;
            //if (obj.GetType().ToString() != "Siemens.Engineering.HmiUnified.UI.Parts.HmiSystemDiagnosisHardwareDetailPart")
            try
            {
                var attrProps = obj.GetAttributes(attrKeys);
                var attrDict = attrKeys.Zip(attrProps, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);


                var colorKeys = from colorAttr in attrDict
                                where colorAttr.Value is Color
                                select colorAttr.Key;
                var colorVals = obj.GetAttributes(colorKeys);
                var colorKeyVals = colorKeys.Zip(colorVals, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
                foreach (var colorKeyVal in colorKeyVals)
                {
                    var color = (Color)colorKeyVal.Value;
                    attrDict.Remove(colorKeyVal.Key);
                    attrDict.Add(colorKeyVal.Key, "0x" + color.A.ToString("X2") + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2"));
                }

                if (parse != null)
                {
                    parse.Invoke(new ParseHelperClass(obj, attrDict));
                    // attrDict = attrDict.Concat(specialDict).ToDictionary(x => x.Key, x => x.Value);
                }

                //Nodes          
                var objKeys = from attributeInfo in obj.GetAttributeInfos()
                              where ((obj.GetAttribute(attributeInfo.Name) as IEngineeringObject) != null
                              && (definedAttributes == null || definedAttributes.Find(definedAttr =>
                              {
                                  var splitAttr = definedAttr.Split('.'); // ["Font", "Name"]
                                  return splitAttr.Take(splitAttr.Count() - 1).Contains(attributeInfo.Name);
                              }) != null))
                              select attributeInfo.Name;

                var objKeyList = objKeys.ToList();
                var objProps = obj.GetAttributes(objKeyList);
                var objDict = objKeyList.Zip(objProps, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);

                foreach (var objKeyVal in objDict)
                {
                    var attr = GetAllMyAttributes(objKeyVal.Value as IEngineeringObject, parse, definedAttributes, fullName + objKeyVal.Key + ".");
                    if (attr.Count != 0)
                    {
                        attrDict.Add(objKeyVal.Key, attr);
                    }
                }

                if (Program.verbose)
                {
                    foreach (var attributeInfo in obj.GetAttributeInfos().Where(attributeInfo => definedAttributes == null || fullName + attributeInfo.Name == "Name" || definedAttributes.Find(x => x == fullName + attributeInfo.Name) != null))
                    {
                        Console.WriteLine(attributeInfo);
                        var test = obj.GetAttribute(attributeInfo.Name);
                        Console.WriteLine(test);
                    }
                }

                if (!(obj is MultilingualText))
                {
                    //Compositions
                    foreach (var compKeyVal in obj.GetCompositionInfos().Where(c => c.Name != "ScreenItems" && c.Name != "PropertyEventHandlers" && c.Name != "Dynamizations" && c.Name != "EventHandlers"
                    && (definedAttributes == null || definedAttributes.Find(definedAttr =>
                    {
                        var splitAttr = definedAttr.Split('.'); // "TrendAreas[0].Trends[0].DataSourceY.Source"  ->  ["TrendAreas[0]", "Trends[0]", "DataSourceY", "Source"]
                        return splitAttr.Take(splitAttr.Count() - 1).FirstOrDefault(x => x.Split('[')[0] == c.Name) != null;
                    }) != null)))
                    {
                        List<object> children = new List<object>();
                        int i = 0;
                        foreach (var item in obj.GetComposition(compKeyVal.Name) as IEngineeringComposition)
                        {
                            children.Add(GetAllMyAttributes(item as IEngineeringObject, parse, definedAttributes, fullName + compKeyVal.Name + "[" + i + "]."));
                            i++;
                        }

                        if (children.Count != 0)
                        {
                            attrDict.Add(compKeyVal.Name, children);
                        }
                    }
                }
                return attrDict;
            }
            catch (Exception ex) 
            { 
                Console.WriteLine("Error with object " + obj.GetType().ToString() + " (" + fullName + "): " + ex.Message);
            }   
            Dictionary<string, object> dic = new Dictionary<string, object>();
            return dic;
        }

        static public void SetAllMyAttributes(Dictionary<string, object> dataTree, IEngineeringObject obj, Func<IEngineeringObject, string, object, IEngineeringObject> reParse)
        {
            foreach (var attr in dataTree)
            {
                if (attr.Value != null && (attr.Value.GetType().Name == "Dictionary`2" || attr.Value.ToString().Contains("List`1"))) //Node
                {
                    if (attr.Key == "Dynamizations" || attr.Key == "Events" || attr.Key == "PropertyEvents")
                    {
                        var specialDict = reParse.Invoke(obj, attr.Key, attr.Value);
                    }
                    else if(attr.Key == "Trigger") 
                    {
                        dataTree = Screen.SetTrigger(dataTree, obj);
                    }
                    else
                    {
                        if (attr.Value.GetType().Name == "Dictionary`2")
                        {
                            var node = obj.GetAttribute(attr.Key) as IEngineeringObject;
                            SetAllMyAttributes((attr.Value as Dictionary<object, object>).ToDictionary(kv => kv.Key.ToString(), kv => kv.Value), node, reParse);
                        }
                        else
                        {
                            SetMyAttributesSimpleTypes(attr.Key, attr.Value, obj);
                        }
                    }
                }
                else //Leave
                {
                    SetMyAttributesSimpleTypes(attr.Key, attr.Value, obj);
                }
            }
        }
        static public void SetMyAttributesSimpleTypes(string keyToSet, object valueToSet, IEngineeringObject obj)
        {
            Type _type = obj.GetType().GetProperty(keyToSet)?.PropertyType;

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
            else if (obj.GetType().Name == "MultilingualText") 
            {
                obj = (obj as MultilingualText).Items.FirstOrDefault(item => item.Language.Culture.Name == keyToSet);

                if(obj == null)
                {
                    Console.WriteLine("Language " + keyToSet + " does not exist in this Runtime!");
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
                if (_type != null)
                {
                    attrVal = Convert.ChangeType(valueToSet, _type);
                } 
            }

            try
            {
                obj.SetAttribute(keyToSet.ToString(), attrVal); 
            }
            catch (Exception ex) { /*Console.WriteLine(ex.Message);*/ } 
        }
        static public void GetMultilingualTextItems(IEngineeringObject obj, ref Dictionary<string, object> multiLingualText)
        {
            foreach (var item in (obj as MultilingualText).Items)
            {
                if (item.Text.ToString() != "")
                {
                    multiLingualText.Add(item.Language.Culture.ToString(), item.Text as object);
                }
            }
        }
        static public bool IsSimple(Type type)
        {
            if(type == null)
            {
                return true;
            }
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                // nullable type, check if the nested type is simple.
                return IsSimple(type.GetGenericArguments()[0]);
            }
            return type.IsPrimitive
              || type.IsEnum
              || type.Equals(typeof(string))
              || type.Equals(typeof(decimal));
        }
    }
}
