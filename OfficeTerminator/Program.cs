using Microsoft.Office.Interop.Word;
using OfficeTerminator.Terminators;
using OfficeTerminator.Terminators.OfficeApplicationTerminator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTerminator
{
    class Program
    {
        static void Main(string[] args)
        {
            new Terminators.OfficeTerminator().TerminateAll();
        }
    }
}
