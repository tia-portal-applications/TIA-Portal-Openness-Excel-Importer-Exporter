using System.Collections.Generic;
using System.Text;
using YamlDotNet.Serialization;
using System.IO;
using System.Linq;
using System;
using System.Globalization;

namespace ExcelExporter
{
    class YAMLConverter : IConverter
    {
        private string folder_old;

        public List<object> Read(string path) //Import
        {
            List<object> data = new List<object>();

            var deserializer = new DeserializerBuilder().Build();

            using (Stream stream = File.OpenRead(path))
            {
                using (TextReader reader = new StreamReader(stream))
                {
                    object result = deserializer.Deserialize<dynamic>(reader);

                    data.Add(result);
                }
            }          
            return data;
        }

        public void Write(Dictionary<string, List<Dictionary<string, object>>> data, string path) //Export
        {           
            var stringBuilder = new StringBuilder();
            var serializer = new SerializerBuilder().DisableAliases().Build();
            stringBuilder.AppendLine(serializer.Serialize(data));

            if (!File.Exists(path))
            {
                using (StreamWriter sw_create = File.CreateText(path))
                {
                    sw_create.Write(stringBuilder.ToString());
                }
            }
            else
            {
                var splitted_file_elements = path.Split('\\');
                var fileNames = path.Split('\\').Last(); //HMIConnection, HMIScreen, ....
                var splitted_file_elements_new = splitted_file_elements.Take(splitted_file_elements.Count() - 1).ToArray();
                var folderName = splitted_file_elements_new.Last(); //temp2
                var splitted_file_elements_new_new = splitted_file_elements_new.Take(splitted_file_elements_new.Count() - 1).ToArray();

                string segment = null;

                string date = String.Format("{0:dd/MM/yyyy H:mm:ss }", Convert.ToDateTime(DateTime.Now), CultureInfo.CreateSpecificCulture("de-DE"));

                string current_date = date.Replace("/", ".").Replace(":", "-");

                foreach (var str in splitted_file_elements_new_new)
                {
                    segment += str.ToString();
                    segment += "\\";
                }


                if (!Directory.Exists(folder_old)) 
                {
                    folder_old = segment + folderName + "oldFiles-" + current_date.Replace("/", ".");
                    DirectoryInfo di = Directory.CreateDirectory(folder_old); //Create new folder
                }

                File.Copy(path, folder_old + "\\" + fileNames, true); //Copy old files from main folder to new folder for documentation

                DirectoryInfo dInfo = new DirectoryInfo(segment + folderName);
                File.Delete(path); //Delete in main folder old files

                using (StreamWriter sw_create = File.CreateText(path))
                {
                    sw_create.Write(stringBuilder.ToString());
                }
            }
        }
    }
}
