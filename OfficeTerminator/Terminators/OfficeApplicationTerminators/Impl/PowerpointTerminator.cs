using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTerminator.Terminators.OfficeApplicationTerminator
{
    public class PowerpointTerminator : OfficeApplicationTerminator<Application>
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
