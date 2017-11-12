//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using Word2007RedactionTool.Properties;

[assembly: CLSCompliant(false)]
namespace Word2007RedactionTool
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			Console.WriteLine("Hello");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //never persist settings x-session
            Settings.Default.Reset();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RedactRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
