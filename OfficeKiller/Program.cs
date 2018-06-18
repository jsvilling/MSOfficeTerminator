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
            OfficeApplicationKiller<Microsoft.Office.Interop.Word.Application> a = new WordKiller();
        }
    }
}
