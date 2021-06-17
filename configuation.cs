using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Newtonsoft.Json;
using System.Reflection;

namespace checker
{
    class Configuation
    {
        public string mou_folder { get; set; }
        public string save_location { get; set; }

        public static Configuation Load() {
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            try {
                using (StreamReader file = new StreamReader(path + "\\config.json"))
                {
                    return JsonConvert.DeserializeObject<Configuation>(file.ReadToEnd());
                }
            } catch (Exception e) {
                return new Configuation();
            }
            
            
        }

        public Configuation(string mou_folder,string save_location) {
            this.mou_folder = mou_folder;
            this.save_location = save_location;
        }

        public Configuation() { }

        public void save() {
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            using (StreamWriter file = new StreamWriter(path+"\\config.json",append:false)) {
                file.WriteLine(JsonConvert.SerializeObject(this));
            }
        }


    }
}
