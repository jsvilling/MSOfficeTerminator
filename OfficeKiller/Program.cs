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
            HashSet<OfficeApplicationKiller> appKillers = new HashSet<OfficeApplicationKiller>();
            appKillers.Add(new ExcelKiller());
            appKillers.Add(new WordKiller());
            appKillers.Add(new PowerpointKiller());
            foreach (OfficeApplicationKiller appKiller in appKillers)
            {
                appKiller.Kill();
                // TODO: error handling
            }
        }
    }
}
