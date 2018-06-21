using Microsoft.Office.Interop.Word;
using OfficeKiller.Killers;
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
            new Killers.OfficeKillerApp().KillAll();
        }
    }
}
