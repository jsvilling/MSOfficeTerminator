using OfficeTerminator.App;
using OfficeTerminator.Terminators.OfficeApplicationTerminator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTerminator.Terminators
{
    public class OfficeTerminator : IOfficeTerminator
    {
        public event ExecutionDoneHandler ExecutionDone;

        public void TerminateAll()
        {
            try
            {
                Properties.Settings.Default.Reload();
                var interfaceType = typeof(IOfficeApplicationTerminator);
                AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(x => x.GetTypes())
                    .Where(x => interfaceType.IsAssignableFrom(x) && !x.IsInterface && !x.IsAbstract)
                    .Select(x => Activator.CreateInstance(x)).ToList()
                    .ForEach(o => o.GetType().GetMethod("Terminate").Invoke(o, null));
                ExecutionDone("All office applications were terminated.");
            }
            catch (Exception e)
            {
                ExecutionDone("An error occured: " + e.ToString());
            }
        }

        public void ChangeConfig(string itemName, bool itemValue)
        {
            if (itemName == "Save files")
            {
                Properties.Settings.Default.SaveFiles = itemValue;
            }
            else if (itemName == "Terminate remaining processes")
            {
                Properties.Settings.Default.TerminateProcess = itemValue;
            }
            Properties.Settings.Default.Save();
        }

        public Dictionary<string, bool> GetConfig()
        {
            Dictionary<string, bool> config = new Dictionary<string, bool>();
            config.Add("Save files", Properties.Settings.Default.SaveFiles);
            config.Add("Terminate remaining processes", Properties.Settings.Default.TerminateProcess);
            return config;
        }
    }
}
