using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace XmlToExcel
{
    public class ArrayNodePathsIdentifier
    {
        public List<List<string>> NodesToFirstDimensionArrays { get; set; } 
        public ArrayNodePathsIdentifier(JObject jObject)
        {
            NodesToFirstDimensionArrays = new List<List<string>>();
            SearchObject(jObject, new List<string>());
        }        
        private void SearchObject(JObject jObject, List<string> nodes)
        {
            var properties = jObject.Properties();
            foreach (var property in properties)
            {
                SearchProperty(property, nodes);
            }
        }

        private void SearchProperty(JProperty property, List<string> nodes)
        {
            var name = property.Name;
            var token = property.Value;

            var localNodes = nodes.ToList();
            localNodes.Add(name);
            if (token is JArray) NodesToFirstDimensionArrays.Add(localNodes);
            if (token is JObject) SearchObject((JObject)token, localNodes);
        }
    }
}