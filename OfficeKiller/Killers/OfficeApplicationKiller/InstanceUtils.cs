using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.Killers.OfficeApplicationKiller
{
    public static class InstanceUtils
    {
        public static T FindRunningInstance<T>(string processName)
        {
            try
            {
                return (T) Marshal.GetActiveObject(processName);
            }
            catch (COMException)
            {
                return default(T);
            }
        }
    }
}
