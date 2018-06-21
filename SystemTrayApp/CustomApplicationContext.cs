using System;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using OfficeKiller.Killers;

namespace SystemTrayApp
{
    public class CustomApplicationContext : ApplicationContext
    {
        public delegate void OnConfigChangedHandler(string itemName, bool itemValue);
        public event OnConfigChangedHandler OnConfigChanged;

        private OfficeAppHandler handler;
        private System.ComponentModel.IContainer components;
        private NotifyIcon notifyIcon;
        private static readonly string DefaultTooltip = "";

        public CustomApplicationContext()
        {
            InitializeContext();
            handler = new OfficeAppHandler();
            handler.ExecutionDone += OnExecutionDone;
            OnConfigChanged += handler.ChangeConfig;
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
                Text = DefaultTooltip,
                Visible = true
            };
            notifyIcon.Text = "Terminate all office programs";
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
            ToolStripMenuItem saveItem = ToolStripMenuItemWithHandler("Save files", OnClickConfigItem);
            ToolStripMenuItem killRemainingItem = ToolStripMenuItemWithHandler("Kill remaining processes", OnClickConfigItem);
            saveItem.Name = "save";
            saveItem.CheckOnClick = true;
            saveItem.Checked = true;
            killRemainingItem.Checked = true;
            killRemainingItem.CheckOnClick = true;
            killRemainingItem.Name = "killProcess";
            menuItem.DropDownItems.Add(saveItem);
            menuItem.DropDownItems.Add(killRemainingItem);
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

        private void notifyIcon_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
                mi.Invoke(notifyIcon, null);
            }
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
