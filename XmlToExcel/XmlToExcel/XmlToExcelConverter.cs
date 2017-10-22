using System;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace XmlToExcel
{
    public class XmlToExcelConverter
    {
        public static void ConvertXmlToExcel(string xml, string outputExcel)
        {
            var completePathOfXmlFile = xml;
            var fileName = outputExcel;

            var jsonObject = GetJsonObject(completePathOfXmlFile);
            var nodesToFirstDimensionArrays = new ArrayNodePathsIdentifier(jsonObject).NodesToFirstDimensionArrays;
            foreach (var nodesToArray in nodesToFirstDimensionArrays)
            {
                var sheetName = GetSheetName(nodesToArray);
                AddExcelSheet(nodesToArray, jsonObject, fileName, sheetName);
                AddDetailedExcelSheet(nodesToArray, jsonObject, fileName, sheetName + "Detailed");
            }
        }


        private static void AddDetailedExcelSheet(List<string> nodesToArray, JObject jsonObject, string fileName, string sheetName)
        {
            var jArray = GetArray(jsonObject, nodesToArray);
            using (var writer = new ExcelWriter(fileName, sheetName))
            {
                writer.WriteHeading("ArrayIndex","Relative Node","Value");
                var nodes = new List<string>();
                for (int i = 0; i < jArray.Count; i++)
                {
                    LogToken(i, nodes, jArray.ElementAt(i), writer);
                }
            }
        }

        private static void LogToken(int arrayIndex, List<string> baseNodes, JToken token, ExcelWriter writer)
        {
            var value = token as JValue;
            if (value != null)
            {
                LogValue(arrayIndex, baseNodes, value, writer);
            }

            var jObject = token as JObject;
            if (jObject != null)
            {
                LogObject(arrayIndex, baseNodes, jObject, writer);
            }

            var jArray = token as JArray;
            if (jArray != null)
            {
                LogArray(arrayIndex, baseNodes, jArray, writer);
            }

        }

        private static void LogValue(int arrayIndex , List<string> baseNodes, JValue value, ExcelWriter writer)
        {
            var settings = new JsonSerializerSettings { DateParseHandling = DateParseHandling.None };
            var container = value.Parent;
            if (container.Type == JTokenType.Property)
            {
                try
                {
                    var data = JsonConvert.DeserializeObject<JObject>("{"+ value.Parent.ToString() +"}", settings).Values()
                        .ElementAt(0);
                    var result = data.Value<string>();
                    LogLine(arrayIndex, baseNodes, result, writer);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }

            }
            else
            {
                LogLine(arrayIndex, baseNodes, value.ToString(), writer);
            }

        }
        private static void LogArray(int arrayIndex, List<string> baseNodes, JArray jArray, ExcelWriter writer)
        {
            for (int i = 0; i < jArray.Count; i++)
            {
                var newNodes = baseNodes.ToList();
                newNodes.Add("["+i+"]");
                LogToken(arrayIndex, newNodes, jArray.ElementAt(i), writer);
            }
        }

        private static void LogObject(int arrayIndex, List<string> baseNodes, JObject jObject, ExcelWriter writer)
        {
            var properties = jObject.Properties();
            foreach (var property in properties)
            {
                var newNodes = baseNodes.ToList();
                newNodes.Add(property.Name);
                LogToken(arrayIndex, newNodes, property.Value, writer);
            }
        }

        private static void LogLine(int arrayIndex, List<string> baseNodes, string value, ExcelWriter writer)
        {
            var stringBuilder = new StringBuilder();
            baseNodes.ForEach(x => stringBuilder.Append(x).Append("\\"));
            if (arrayIndex % 2 == 0)
            {
                writer.WriteLineBlue(arrayIndex.ToString(), stringBuilder.ToString(), value);
            }
            else
            {
                writer.WriteLineGreen(arrayIndex.ToString(), stringBuilder.ToString(), value);
            }
        }

        private static JObject GetJsonObject(string completePathOfXmlFile)
        {
            var xml = new XmlDocument();
            xml.Load(completePathOfXmlFile);
            string json = JsonConvert.SerializeXmlNode(xml);
            var deserializedObject = JsonConvert.DeserializeObject(json);
            var formatedJSonString = JsonConvert.SerializeObject(deserializedObject, Newtonsoft.Json.Formatting.Indented);
            JObject jsonObject = JObject.Parse(formatedJSonString);
            return jsonObject;
        }

        private static void AddExcelSheet(List<string> nodesToArray, JObject jsonObject, string fileName, string sheetName)
        {
            var jArray = GetArray(jsonObject, nodesToArray);
            Dictionary<string, int> dictionary = new Dictionary<string, int>();
            int entryIndex = 1;
            bool hasSimpleValue = false;
            foreach (var element in jArray)
            {
                if (element is JObject)
                {
                    var properties = ((JObject)element).Properties();
                    foreach (var property in properties)
                    {
                        if (!dictionary.ContainsKey(property.Name))
                        {
                            dictionary.Add(property.Name, entryIndex);
                            entryIndex++;
                        }
                    }
                }
                if (element is JValue)
                {
                    hasSimpleValue = true;
                }
            }

            using (var writer = new ExcelWriter(fileName, sheetName))
            {
                var keys = dictionary.Keys.ToList();
                if(hasSimpleValue) keys.Insert(0, "Value");
                writer.WriteHeading(keys.ToArray());


                foreach (var element in jArray)
                {
                    if (element is JObject)
                    {
                        List<string> results = new List<string>();
                        if (hasSimpleValue) results.Insert(0,string.Empty);
                        foreach (var key in keys)
                        {
                            var p = ((JObject)element).Property(key);
                            if (p != null)
                            {
                                var text = p.Value.ToString();
                                if (p.Value is JValue)
                                {
                                    var newlyReadText = GetValueToPreventInterpretationOfDate(p.ToString());
                                    if (!string.IsNullOrEmpty(newlyReadText))
                                    {
                                        text = newlyReadText;
                                    }

                                }
                                text = text.Replace(Environment.NewLine, string.Empty);
                                text = Regex.Replace(text, @"[^\S\r\n]+", " ");
                                results.Add(text);
                            }
                            else
                            {
                                results.Add(string.Empty);
                            }
                        }
                        writer.Write(results.ToArray());
                    }
                    if (element is JValue)
                    {
                        writer.Write(element.ToString());
                    }
                }
            }
        }

        private static string GetValueToPreventInterpretationOfDate(string input)
        {
            try
            {
                var settings = new JsonSerializerSettings { DateParseHandling = DateParseHandling.None };
                var data = JsonConvert.DeserializeObject<JObject>("{" + input + "}", settings).Values().ElementAt(0);
                return data.Value<string>();
            }
            catch (Exception e)
            {
                return String.Empty;
            }
        }

        private static string GetSheetName(List<string> nodesToArray)
        {
            var sheetName = nodesToArray.LastOrDefault();
            if (string.IsNullOrEmpty(sheetName)) sheetName = "Default";
            Regex regex = new Regex("[^a-zA-Z0-9 -]");
            sheetName = regex.Replace(sheetName, "");
            return sheetName;
        }


        public static JArray GetArray(JObject jObject, List<string> paths)
        {
            var last = paths.Last();
            var newPaths = paths.ToList();
            newPaths.RemoveAt(newPaths.Count - 1);
            var resultJObject = GetObject(jObject, newPaths);
            return (JArray)resultJObject[last];
        }

        public static JObject GetObject(JObject jObject, List<string> paths)
        {
            if (paths.Count == 0) return jObject;
            var resultObject = (JObject)(jObject[paths.ElementAt(0)]);
            var newList = paths.ToList();
            newList.RemoveAt(0);
            return GetObject(resultObject, newList);
        }
    }
}
