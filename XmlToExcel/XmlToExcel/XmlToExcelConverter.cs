﻿using System;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
            //AddCompleteSummary(jsonObject, fileName, "CompleteSummary");
            var nodesToFirstDimensionArrays = new ArrayNodePathsIdentifier(jsonObject).NodesToFirstDimensionArrays;
            foreach (var nodesToArray in nodesToFirstDimensionArrays)
            {
                var sheetName = GetSheetName(nodesToArray);
                AddExcelSheet(nodesToArray, jsonObject, fileName, sheetName);
                AddDetailedExcelSheet(nodesToArray, jsonObject, fileName, sheetName + "Detailed");
            }
        }

        private static void AddCompleteSummary(JObject jsonObject, string fileName, string sheetName)
        {
            using (var writer = new ExcelWriter(fileName, sheetName))
            {
                writer.WriteHeading("Path to Node", "Value");
                Action<string, string> writeFunction = (a, b) =>
                {
                    writer.Write( a, b);
                };
                LogObject(0, new List<string>(),jsonObject,writeFunction );
            }
        }


        private static void AddDetailedExcelSheet(List<string> nodesToArray, JObject jsonObject, string fileName, string sheetName)
        {
            var jArray = GetArray(jsonObject, nodesToArray);
            using (var writer = new ExcelWriter(fileName, sheetName))
            {
                writer.WriteHeading("ArrayIndex","Relative Node","Value");
                var detailedReportWriter = new DetailedReportWriter(writer);

                var nodes = new List<string>();
                for (int i = 0; i < jArray.Count; i++)
                {
                    Dictionary<string, string> rowData = new Dictionary<string, string>();
                    Action<string,string> writeFunction = (a,b) =>
                    {
                        rowData.Add(a,b);
                    };
                    LogToken(i, nodes, jArray.ElementAt(i), writeFunction);
                    detailedReportWriter.Write(rowData);
                }
                detailedReportWriter.Dispose();
            }
        }

        private static void LogToken(int arrayIndex, List<string> baseNodes, JToken token, Action<string,string> writer)
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

        private static void LogValue(int arrayIndex , List<string> baseNodes, JValue value, Action<string, string> writer)
        {
            
            var result = GetStringValue(value);
            LogLine(arrayIndex, baseNodes, result, writer);

        }

        private static string GetStringValue(JValue value)
        {
            var settings = new JsonSerializerSettings { DateParseHandling = DateParseHandling.None };
            var container = value.Parent;
            if (container.Type == JTokenType.Property)
            {
                var data = JsonConvert.DeserializeObject<JObject>("{" + value.Parent.ToString() + "}", settings).Values()
                    .ElementAt(0);
                var result = data.Value<string>();
                if (string.IsNullOrWhiteSpace(result)) return "";
                return result;
            }
            else
            {
                return value.ToString();
            }
        }

        private static void LogArray(int arrayIndex, List<string> baseNodes, JArray jArray, Action<string, string> writer)
        {
            for (int i = 0; i < jArray.Count; i++)
            {
                var newNodes = baseNodes.ToList();
                newNodes.Add("["+i+"]");
                LogToken(arrayIndex, newNodes, jArray.ElementAt(i), writer);
            }
        }

        private static void LogObject(int arrayIndex, List<string> baseNodes, JObject jObject, Action<string, string> writer)
        {
            var properties = jObject.Properties();
            foreach (var property in properties)
            {
                var newNodes = baseNodes.ToList();
                newNodes.Add(property.Name);
                LogToken(arrayIndex, newNodes, property.Value, writer);
            }
        }

        private static void LogLine(int arrayIndex, List<string> baseNodes, string value, Action<string, string> writer)
        {
            var stringBuilder = new StringBuilder();
            baseNodes.ForEach(x => stringBuilder.Append(x).Append("\\"));
            writer.Invoke(stringBuilder.ToString(), value);
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
                                    text = GetStringValue((JValue)p.Value);
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



