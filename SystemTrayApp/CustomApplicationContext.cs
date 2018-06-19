using System;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using OfficeKiller.Killers;

namespace SystemTrayApp
{
    public class CustomApplicationContext : ApplicationContext
    {
        OfficeAppHandler handler;
        private System.ComponentModel.IContainer components;
        private NotifyIcon notifyIcon;
        private static readonly string IconFileName = "icon.ico";
        private static readonly string DefaultTooltip = "";

        public CustomApplicationContext()
        {
            InitializeContext();
            handler = new OfficeAppHandler();
        }

        private void InitializeContext()
        {
            components = new System.ComponentModel.Container();
            notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = new Icon(IconFileName),
                Text = DefaultTooltip,
                Visible = true
            };
            notifyIcon.ContextMenuStrip.Items.Add(ToolStripMenuItemWithHandler("terminate all office programs", OnTerminateAllOfficeApps));
            notifyIcon.DoubleClick += OnTerminateAllOfficeApps;
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

        private void OnTerminateAllOfficeApps(object sender, EventArgs e)
        {
            handler.KillAll();
        }
    }
}
