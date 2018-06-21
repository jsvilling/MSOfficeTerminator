using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTerminator.Terminators.OfficeApplicationTerminator
{
    public class WordTerminator : OfficeApplicationTerminator<Application>
    {
        public override string InstanceName => "Word.Application";

        public override string ProcessName => "WINWORD";

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
