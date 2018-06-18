using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.Killers.OfficeApplicationKiller
{
    public abstract class OfficeApplicationKiller<A>
    {
        public abstract string InstanceName
        {
            get;
        }

        public abstract string ProcessName
        {
            get;
        }

        public void Kill()
        {
            SaveAndQuitRunningApplications();
            KillRemainingProcesses();
        }

        public A FindRunningInstance()
        {
            try
            {
                return (A) Marshal.GetActiveObject(InstanceName);
            }
            catch (COMException)
            {
                return default(A);
            }
        }

        private void SaveAndQuitRunningApplications()
        {
            A appInstance = FindRunningInstance();
            if (appInstance != null)
            {
                SaveAndQuit(appInstance);
            }
        }

        private void KillRemainingProcesses()
        {
            Process[] remainingProcesses = Process.GetProcessesByName(ProcessName);
            foreach (Process process in remainingProcesses)
            {
                process.Kill();
            }
        }

        protected abstract void SaveAndQuit(A appInstance);
    }
}
