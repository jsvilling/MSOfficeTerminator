using System;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using OfficeKiller.Killers;
using System.Collections.Generic;
using OfficeKiller.App;

namespace SystemTrayApp
{
    public class CustomApplicationContext : ApplicationContext
    {
        public delegate void OnConfigChangedHandler(string itemName, bool itemValue);
        public event OnConfigChangedHandler OnConfigChanged;

        private IOfficeKillerApp handler;
        private System.ComponentModel.IContainer components;
        private NotifyIcon notifyIcon;

        public CustomApplicationContext(IOfficeKillerApp handler)
        {
            this.handler = new OfficeKillerApp();
            handler.ExecutionDone += OnExecutionDone;
            OnConfigChanged += handler.ChangeConfig;
            InitializeContext();
        }

        private void InitializeContext()
        {
            components = new System.ComponentModel.Container();
            notifyIcon = SetupNotifyIcon();
        }

        private NotifyIcon SetupNotifyIcon()
        {
            NotifyIcon notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = Properties.Resources.TrayIcon,
                Text = "Terminate all office programs",
                Visible = true
            };
            notifyIcon.ContextMenuStrip.Items.Add(SetupConfigSubMenu());
            notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            notifyIcon.ContextMenuStrip.Items.Add(ToolStripMenuItemWithHandler("Terminate all office programs", OnTerminateAllOfficeApps));
            notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            notifyIcon.ContextMenuStrip.Items.Add(ToolStripMenuItemWithHandler("Exit", OnExit));
            notifyIcon.DoubleClick += OnTerminateAllOfficeApps;
            return notifyIcon;
        }

        private ToolStripMenuItem SetupConfigSubMenu()
        {
            ToolStripMenuItem menuItem = new ToolStripMenuItem("Configuration");
            foreach (KeyValuePair<string, bool> configItem in handler.GetConfig())
            {
                ToolStripMenuItem item = ToolStripMenuItemWithHandler(configItem.Key, OnClickConfigItem);
                item.Name = configItem.Key;
                item.Checked = configItem.Value;
                item.CheckOnClick = true;
                menuItem.DropDownItems.Add(item);
            }
            return menuItem;
        }

        private ToolStripMenuItem ToolStripMenuItemWithHandler(string displayText, EventHandler eventHandler)
        {
            var item = new ToolStripMenuItem(displayText);
            if (eventHandler != null)
            {
                item.Click += eventHandler;
            }
            return item;
        }

        // Event Handlers

        private void OnExecutionDone(string message)
        {
            notifyIcon.BalloonTipText = message;
            notifyIcon.ShowBalloonTip(2000);
        }

        private void OnClickConfigItem(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (ToolStripMenuItem) sender;
            OnConfigChanged(item.Name, item.Checked);
        }

        private void OnTerminateAllOfficeApps(object sender, EventArgs e)
        {
            handler.KillAll();
        }

        private void OnExit(object sender, EventArgs e)
        {
            this.ExitThread();
        }
    }
}