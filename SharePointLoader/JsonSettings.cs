using Newtonsoft.Json;
 
using System; 
using System.IO; 

namespace SharePointLoader
{


    public static class JsonSettings
    {

        private static   string FolderPath => System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
        public static void Save<T>(T settings) where T : class
        {
            string path = Path.Combine(FolderPath, typeof(T).ToString() + ".json");
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(path))
                {
                    streamWriter.WriteLine(JsonConvert.SerializeObject((object)settings, Formatting.Indented));
                    streamWriter.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static T Get<T>() where T : class
        {
            string path =Path.Combine(FolderPath, typeof(T).ToString() + ".json");
            if (File.Exists(path))
            {
                using (StreamReader streamReader = new StreamReader(path))
                {
                    T result = JsonConvert.DeserializeObject<T>(streamReader.ReadToEnd());
                    streamReader.Close();  
                    return result;
                }
            }
            else
            {
                throw new FileNotFoundException("File with configurations not found", path);
            } 
        }

    }
}
