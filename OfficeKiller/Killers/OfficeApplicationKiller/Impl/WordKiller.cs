using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.Killers.OfficeApplicationKiller
{
    public class WordKiller : OfficeApplicationKiller<Application>
    {
        public override string InstanceName => "Word.Application";

        public override string ProcessName => "winword.exe";

        protected override void SaveAndQuit(Application runningWordApp)
        {
            foreach (Document wordDoc in runningWordApp.Documents)
            {
                wordDoc.Save();
                wordDoc.Close();
            }
            runningWordApp.Quit();
        }
    }
}
