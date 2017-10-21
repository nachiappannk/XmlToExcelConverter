using System;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
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
                AddExcelSheet(nodesToArray, jsonObject, fileName);
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

        private static void AddExcelSheet(List<string> nodesToArray, JObject jsonObject, string fileName)
        {
            var sheetName = GetSheetName(nodesToArray);

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
                                //text = text.Replace(Environment.NewLine, string.Empty);
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
