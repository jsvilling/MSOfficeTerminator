using OfficeKiller.App;
using OfficeKiller.Killers.OfficeApplicationKiller;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.Killers
{
    public class OfficeAppHandler
    {
        public delegate void ExecutionDoneHandler(string message);
        public event ExecutionDoneHandler ExecutionDone;
        string message = "";

        public void KillAll()
        {
            try
            {
                var interfaceType = typeof(IOfficeApplicationKiller);
                AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(x => x.GetTypes())
                    .Where(x => interfaceType.IsAssignableFrom(x) && !x.IsInterface && !x.IsAbstract)
                    .Select(x => Activator.CreateInstance(x)).ToList()
                    .ForEach(o => o.GetType().GetMethod("Kill").Invoke(o, null));
                message = "All office applications were terminated.";
            }
            catch (Exception e)
            {
                message = "An error occured: " + e.ToString();
            }
            finally
            {
                ExecutionDone(message);
                message = "";
            }
        }

        public void ChangeConfig()
        {

        }
    }
}
