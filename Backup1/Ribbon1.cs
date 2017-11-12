// Copyright (c) Microsoft Corporation.  All rights reserved.
using Office = Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

using Word2007RedactionTool.Properties;

namespace Word2007RedactionTool
{
    [ComVisible(true)]
    public partial class RedactRibbon : Office.IRibbonExtensibility
    {      
        private static Word.WdColor ShadingColor = (Word.WdColor)12697792; //marks are gray by default, but can be changed by setting this property.

        private object Missing = Type.Missing;
        private object CollapseStart = Word.WdCollapseDirection.wdCollapseStart;
        private object CollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;

        private Office.IRibbonUI Ribbon;
        private Word.Application Application;
        private Word.ApplicationEvents4_WindowSelectionChangeEventHandler SelectionChangeEvent;
        private Word.ApplicationEvents4_DocumentChangeEventHandler DocumentChangeEvent;      

        public RedactRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {
            return GetResourceText("Word2007RedactionTool.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            //hold a handle to the ribbon so we can refresh controls later
            this.Ribbon = ribbonUI;

            //register events
            Application = Globals.ThisAddIn.Application;
            SelectionChangeEvent = new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            Application.WindowSelectionChange += SelectionChangeEvent;
            DocumentChangeEvent = new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            Application.DocumentChange += DocumentChangeEvent;
        }

        public System.Drawing.Bitmap Ribbon_LoadImages(string image)
        {
            return GetResourceImage(image);
        }

        public bool Ribbon_GetEnabled(Office.IRibbonControl control)
        {
            if (Application.Documents.Count == 0)
                return false;
            else if ((control.Id == "splitButtonMark" || control.Id == "splitButtonUnmark") && Application.Selection != null && Application.Selection.Type == Word.WdSelectionType.wdSelectionColumn)
                return false;
            else if (control.Id != "buttonMarkOfficeMenu" && Application.Selection != null && Application.Selection.StoryType == Word.WdStoryType.wdCommentsStory)
                return false;
            else if (Application.ActiveDocument.ProtectionType != Word.WdProtectionType.wdNoProtection)
                return false;
            else
                return true;
        }

        public string Ribbon_GetLabel(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "buttonMarkOfficeMenu":    
                case "buttonRedact":
                    return Resources.RedactMenuItemLabel;
                case "groupRedact":
                    return Resources.RedactGroupLabel;
                case "buttonUnmark":
                case "splitButtonUnmark__btn":
                    return Resources.UnmarkLabel;
                case "buttonUnmarkAll":
                    return Resources.UnmarkAllLabel;
                case "buttonPrevious":
                    return Resources.PreviousLabel;
                case "buttonNext":
                    return Resources.NextLabel;
                case "buttonMark":
                case "splitButtonMark__btn":
                    return Resources.MarkLabel;
                case "buttonFindAndMark":
                    return Resources.FindAndMarkLabel;
                default:
                    Debug.Fail("unknown control requested a label: " + control.Id);
                    return null;
            }
        }

        public string Ribbon_GetDescription(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "buttonMarkOfficeMenu":
                    return Resources.RedactMenuItemDescription;
                default:
                    Debug.Fail("unknown control requested a description: " + control.Id);
                    return null;
            }
        }

        public string Ribbon_GetScreentip(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButtonMark__btn":
                case "buttonMark":
                    return Resources.MarkScreentip;
                case "buttonRedact":
                    return Resources.RedactScreentip;
                case "splitButtonUnmark__btn":
                case "buttonUnmark":
                    return Resources.UnmarkScreentip;
                case "buttonPrevious":
                    return Resources.PreviousScreentip;
                case "buttonNext":
                    return Resources.NextScreentip;
                default:
                    Debug.Fail("unknown control requested a screentip: " + control.Id);
                    return null;
            }
        }

        public string Ribbon_GetSupertip(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButtonMark__btn":
                case "buttonMark":
                    return Resources.MarkSupertip;
                case "splitButtonMark":
                    return Resources.MarkSplitMenuSupertip;
                case "buttonRedact":
                    return Resources.RedactSupertip;
                case "splitButtonUnmark__btn":
                case "buttonUnmark":
                    return Resources.UnmarkSupertip;
                case "splitButtonUnmark":
                    return Resources.UnmarkSplitMenuSupertip;
                case "buttonPrevious":
                    return Resources.PreviousSupertip;
                case "buttonNext":
                    return Resources.NextSupertip;
                case "buttonFindAndMark":
                    return Resources.FindAndMarkSupertip;
                default:
                    Debug.Fail("unknown control requested a supertip: " + control.Id);
                    return null;
            }
        }

        public string Ribbon_GetKeytip(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "splitButtonMark":
                    return Resources.MarkKeytip;
                case "splitButtonUnmark":
                    return Resources.UnmarkKeytip;
                case "buttonPrevious":
                    return Resources.PreviousKeytip;
                case "buttonNext":
                    return Resources.NextKeytip;
                default:
                    Debug.Fail("unknown control requested a supertip: " + control.Id);
                    return null;
            }
        }

        public void SplitButtonMark_Click(Office.IRibbonControl control)
        {
            TurnOffEvents();
            MarkSelection();
            TurnOnEvents();
        }

        public void ButtonUnmark_Click(Office.IRibbonControl control)
        {
            TurnOffEvents();
            UnmarkSelection();
            TurnOnEvents();
        }

        public void ButtonUnmarkAll_Click(Office.IRibbonControl control)
        {
            TurnOffEvents();
            UnmarkDocument();
            TurnOnEvents();
        }
       
        public void ButtonPrevious_Click(Office.IRibbonControl control)
        {
            TurnOffEvents();
            SelectPreviousMark();
            TurnOnEvents();
        }        

        public void ButtonNext_Click(Office.IRibbonControl control)
        {
            TurnOffEvents();
            SelectNextMark();
            TurnOnEvents();
        }

        public void ButtonRedact_Click(Office.IRibbonControl control)
        {
            TurnOffEvents();
            RedactDocument();
            TurnOnEvents();
        }

        public void ButtonFindAndMark_Click(Office.IRibbonControl control)
        {
            TurnOffEvents();
            FindAndMark();
            TurnOnEvents();
        }

        #endregion

        #region Word Events

        private void Application_WindowSelectionChange(Word.Selection Selection)
        {
            InvalidateRedactionControls();
        }

        private void Application_DocumentChange()
        {
            InvalidateRedactionControls();
        }

        /// <summary>
        /// Invalidate the controls added by the redaction tool.
        /// </summary>
        private void InvalidateRedactionControls()
        {
            Ribbon.InvalidateControl("splitButtonMark");
            Ribbon.InvalidateControl("splitButtonUnmark");
            Ribbon.InvalidateControl("buttonPrevious");
            Ribbon.InvalidateControl("buttonNext");
        }

        private void TurnOffEvents()
        {
            try
            {
                Application.WindowSelectionChange -= SelectionChangeEvent;
                Application.DocumentChange -= DocumentChangeEvent;
            }
            catch (NullReferenceException) // if we fail to finish something, turn on never gets called, which would cause this to fail forever
            { }
        }

        private void TurnOnEvents()
        {
            Application.WindowSelectionChange += SelectionChangeEvent;
            Application.DocumentChange += DocumentChangeEvent;
            InvalidateRedactionControls();
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            Debug.Fail("Failed to get the ribbon xml.");
            return null;
        }

        private static System.Drawing.Bitmap GetResourceImage(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (resourceNames[i].EndsWith(resourceName, StringComparison.OrdinalIgnoreCase))
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            Icon iconImage = new Icon(resourceReader.BaseStream);
                            return iconImage.ToBitmap();
                        }
                    }
                }
            }
            Debug.Fail("Failed to get the ribbon icon.");
            return null;
        }

        #endregion
    }
}
