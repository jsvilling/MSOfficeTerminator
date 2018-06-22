using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTerminator.App.Configuration
{
    public class AppConfig
    {
        private static Dictionary<string, bool> config;

        public static KeyValuePair<string, bool> Get(string name)
        {
            return new KeyValuePair<string, bool>(name, config[name]);
        }

        private static void Init()
        {
            System.Collections.IEnumerator settingsEnumerator = Properties.Settings.Default.Properties.GetEnumerator();
            while (settingsEnumerator.MoveNext())
            {

                System.Configuration.SettingsProperty setting = settingsEnumerator.Current as System.Configuration.SettingsProperty;
            }
        }

    }

    public class ConfigItem
    {
        string name;
        bool value;

        public ConfigItem(string name, bool value)
        {
            this.name = name;
            this.value = value;
        }
    }
}
