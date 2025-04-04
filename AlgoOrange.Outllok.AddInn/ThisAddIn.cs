using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace AlgoOrange.Outllok.AddInn
{
    public partial class ThisAddIn
    {
        private ChatControl chatControl;
        private Microsoft.Office.Tools.CustomTaskPane chatTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            chatControl = new ChatControl();
            chatTaskPane = this.CustomTaskPanes.Add(chatControl, "Chat");
            chatTaskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void ToggleChatControlVisibility()
        {
            chatTaskPane.Visible = !chatTaskPane.Visible;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
