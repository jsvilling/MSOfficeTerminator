using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.Killers.OfficeApplicationKiller
{
    public class WordKiller : OfficeApplicationKiller
    {
        public void Kill()
        {
            Application runningWordApp = FindRunningWordInstance();
            if (runningWordApp != null)
            {
                KillWord(runningWordApp);
            }
        }

        private Application FindRunningWordInstance()
        {
            return InstanceUtils.FindRunningInstance<Application>("Word.Application");
        }

        private void KillWord(Application runningWordApp)
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
