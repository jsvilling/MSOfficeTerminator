﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTerminator.Terminators.OfficeApplicationTerminator
{
    public abstract class OfficeApplicationTerminator<A> : IOfficeApplicationTerminator
    {
        public abstract string InstanceName
        {
            get;
        }

        public abstract string ProcessName
        {
            get;
        }

        public void Terminate()
        {
            if (Properties.Settings.Default.SaveFiles)
            {
                SaveAndQuitRunningApplications();
            }

            if (Properties.Settings.Default.TerminateProcess)
            {
                TerminateRemainingProcesses();
            }
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

        private void TerminateRemainingProcesses()
        {
            Process.GetProcessesByName(ProcessName).ToList().ForEach(p => p.Kill());
        }

        protected abstract void SaveAndQuit(A appInstance);
    }
}
