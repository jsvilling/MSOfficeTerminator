using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.Killers.OfficeApplicationKiller
{
    public class PowerpointKiller : OfficeApplicationKiller<Application>
    {
        public override string InstanceName => "PowerPoint.Application";

        public override string ProcessName => "POWERPNT";

        protected override void SaveAndQuit(Application appInstance)
        {
            foreach (Presentation presentation in appInstance.Presentations)
            {
                presentation.Save();
                presentation.Close();
            }
            appInstance.Quit();
        }
    }
}
