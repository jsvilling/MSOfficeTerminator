using Microsoft.Office.Interop.Word;
using OfficeKiller.Killers.OfficeApplicationKiller;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller
{
    class Program
    {
        static void Main(string[] args)
        {
            var interfaceType = typeof(IOfficeApplicationKiller);
            AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(x => x.GetTypes())
                .Where(x => interfaceType.IsAssignableFrom(x) && !x.IsInterface && !x.IsAbstract)
                .Select(x => Activator.CreateInstance(x)).ToList()
                .ForEach(o => o.GetType().GetMethod("Kill").Invoke(o, null));
        }
    }
}
