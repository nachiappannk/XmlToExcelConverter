using System.Collections.Generic;

namespace XmlToExcel
{
    public class DetailedReportWriter
    {
        private readonly ExcelWriter _writer;
        private Dictionary<string, int> _columnIndexDictionary;
        private int _numberOfColumns = 0;

        public DetailedReportWriter(ExcelWriter writer)
        {
            _writer = writer;
            _columnIndexDictionary = new Dictionary<string, int>();
        }

        public void Write(Dictionary<string, string> dictionary)
        {
            foreach (var keyValuePair in dictionary)
            {
                var key = keyValuePair.Key;
                if (!IsColumnNameAvailable(key)) AddNewColumnName(key);
            }

            string[] objectsToWrite = new string[_numberOfColumns];
            for (int i = 0; i < _numberOfColumns; i++)
            {
                objectsToWrite[i] = string.Empty;
            }


            foreach (var keyValuePair in dictionary)
            {
                var key = keyValuePair.Key;
                int index = _columnIndexDictionary[key];
                objectsToWrite[index - 1] = keyValuePair.Value;
            }
            _writer.Write(objectsToWrite);
        }

        private bool IsColumnNameAvailable(string key)
        {
            return _columnIndexDictionary.ContainsKey(key);
        }

        private void AddNewColumnName(string key)
        {
            _numberOfColumns++;
            _columnIndexDictionary.Add(key, _numberOfColumns);
        }

        public void Dispose()
        {
            string[] headingsToWrite = new string[_numberOfColumns];
            for (int i = 0; i < _numberOfColumns; i++)
            {
                headingsToWrite[i] = string.Empty;
            }
            foreach (var keyValuePair in _columnIndexDictionary)
            {
                headingsToWrite[keyValuePair.Value - 1] = keyValuePair.Key;
            }
            _writer.WriteHeading(headingsToWrite);
        }
    }
}