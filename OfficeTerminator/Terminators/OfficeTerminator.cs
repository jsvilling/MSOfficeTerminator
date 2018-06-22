﻿using OfficeTerminator.App;
using OfficeTerminator.Terminators.OfficeApplicationTerminator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTerminator.Terminators
{
    public class OfficeTerminator : IOfficeTerminator
    {
        public event ExecutionDoneHandler AllExecutionDone;
        private int count = 0;

        public void TerminateAll()
        {
            try
            {
                Properties.Settings.Default.Reload();
                foreach (IOfficeApplicationTerminator terminator in GetTerminators())
                {
                    count++;
                    ThreadWorker tw = new ThreadWorker(terminator);
                    tw.TerminatorDone += OnTerminatorDone;
                    new Thread(tw.Run).Start();
                }
            }
            catch (Exception e)
            {
                AllExecutionDone("An error occured: " + e.ToString());
            }
        }

        private void OnTerminatorDone(object sender, EventArgs e)
        {
            if (count-- == 1)
            {
                AllExecutionDone("All office applications were terminated.");
            }
        }

        class ThreadWorker
        {
            public event EventHandler TerminatorDone;
            private IOfficeApplicationTerminator terminator;
            public ThreadWorker(IOfficeApplicationTerminator terminator)
            {
                this.terminator = terminator;
            }

            public void Run()
            {
                terminator.Terminate();
                TerminatorDone(this, EventArgs.Empty);
            }
        }

        private List<IOfficeApplicationTerminator> GetTerminators()
        {
            var interfaceType = typeof(IOfficeApplicationTerminator);
            return AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(x => x.GetTypes())
                .Where(x => interfaceType.IsAssignableFrom(x) && !x.IsInterface && !x.IsAbstract)
                .Select(x => Activator.CreateInstance(x) as IOfficeApplicationTerminator).ToList();
        }

        public void ChangeConfig(string itemName, bool itemValue)
        {
            if (itemName == "Save files")
            {
                Properties.Settings.Default.SaveFiles = itemValue;
            }
            else if (itemName == "Terminate remaining processes")
            {
                Properties.Settings.Default.TerminateProcess = itemValue;
            }
            Properties.Settings.Default.Save();
        }

        public Dictionary<string, bool> GetConfig()
        {
            Dictionary<string, bool> config = new Dictionary<string, bool>();
            config.Add("Save files", Properties.Settings.Default.SaveFiles);
            config.Add("Terminate remaining processes", Properties.Settings.Default.TerminateProcess);
            return config;
        }
    }
}
