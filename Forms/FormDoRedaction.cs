//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Globalization;

namespace Word2007RedactionTool
{
    public partial class FormDoRedaction : Form
    {
        private object False = false;
        private object Missing = Type.Missing;

        private Word.Document FileToRedact;
        private RedactRibbon RedactRibbon;        
        private System.ComponentModel.BackgroundWorker backgroundWorkerRedact;

        public FormDoRedaction(Word.Document document, RedactRibbon ribbon)
        {
            InitializeComponent();

            FileToRedact = document;
            RedactRibbon = ribbon;

            //progress = 100(main doc) + 17 (the # of secondary stories)
            progressBar.Maximum = 117;

            //set up a background worker to do the redaction (so we can report progress)
            this.backgroundWorkerRedact = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorkerRedact.WorkerReportsProgress = true;
            this.backgroundWorkerRedact.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerRedact_DoWork);
            this.backgroundWorkerRedact.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerRedact_RunWorkerCompleted);
            this.backgroundWorkerRedact.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorkerRedact_ProgressChanged);
        }

        private void FormDoRedaction_Load(object sender, EventArgs e)
        {
            //wait cursor
            FileToRedact.Application.System.Cursor = Word.WdCursorType.wdCursorWait;

            //redact on that another thread
            backgroundWorkerRedact.RunWorkerAsync(FileToRedact);

            //for debugging only
            //bool Succeeded = RedactRibbon.RedactDocument(FileToRedact, null);
            //DoPostRedactionTasks(new RedactResult(FileToRedact, Succeeded));
        }

        private void backgroundWorkerRedact_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {            
            Word.Document Document = (Word.Document)e.Argument;

            //redact the document
            bool Succeeded = RedactRibbon.RedactDocument(Document, backgroundWorkerRedact);
            FillProgressBar();  
            
            //return the result
            e.Result = new RedactResult(Document, Succeeded);
        }

        private void backgroundWorkerRedact_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            DoPostRedactionTasks((RedactResult)e.Result);
        }

        private void backgroundWorkerRedact_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            UpdateProgress(e.ProgressPercentage);
        }

        /// <summary>
        /// Performs tasks after a redaction has been completed (succesfully or not).
        /// </summary>
        /// <param name="Result">Specifies the results of the redaction attempt.</param>
        private void DoPostRedactionTasks(RedactResult Result)
        {
            Word.Document Document = Result.Document;
            Word.Application Application = Document.Application;

            if (Result.Succeeded)
            {
                //post redaction tasks
                Document.Fields.Update(); //make sure any reference fields have the redacted results
                Document.Application.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView;
                Document.Saved = true;

                //finally, blast the undo stack
                Document.UndoClear();

                Document.Application.ScreenUpdating = true;
                Document.Application.ScreenRefresh();

                this.DialogResult = DialogResult.OK;
            }
            else
            {
                Document.Application.ScreenUpdating = true;
                Document.Application.ScreenRefresh();

                ((Word._Document)Document).Close(ref False, ref Missing, ref Missing);

                this.DialogResult = DialogResult.Abort;
            }

            Application.System.Cursor = Word.WdCursorType.wdCursorNormal;

            //close the dialog
            CloseDialog();
        }

        #region UI methods

        internal delegate void ResetProgressDelegate();
        internal void ResetProgress()
        {
            if (InvokeRequired)
            {
                Invoke(new ResetProgressDelegate(ResetProgress), new object[] { });
                return;
            }

            progressBar.Value = 0;
        }

        internal delegate void CloseDialogDelegate();
        internal void CloseDialog()
        {
            if (InvokeRequired)
            {
                Invoke(new CloseDialogDelegate(CloseDialog), new object[] { });
                return;
            }

            this.Close();
        }

        internal delegate void UpdateProgressDelegate(int Percentage);
        internal void UpdateProgress(int Percentage)
        {
            if (InvokeRequired)
            {
                Invoke(new UpdateProgressDelegate(UpdateProgress), new object[] { Percentage });
                return;
            }
            
            if(Percentage > progressBar.Value)
                progressBar.Value = Percentage;

            UpdateProgressBarText(Percentage.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Debug only. Shows the progress bar value in the text label.
        /// </summary>
        /// <param name="labelText">The value of the progress bar.</param>
        [Conditional("DEBUG")]
        private void UpdateProgressBarText(string labelText)
        {
            labelProgress.Text = labelText;
        }

        internal delegate void FillProgressBarDelegate();
        internal void FillProgressBar()
        {
            if (InvokeRequired)
            {
                Invoke(new FillProgressBarDelegate(FillProgressBar), new object[] { });
                return;
            }

            progressBar.Value = progressBar.Maximum;
        }

        #endregion
    }

    /// <summary>
    /// Specifies the results of a redaction operation.
    /// </summary>
    struct RedactResult
    {
        public readonly Word.Document Document;
        public readonly bool Succeeded;

        public RedactResult(Word.Document Doc, bool Success)
        {
            Document = Doc;
            Succeeded = Success;
        }
    }
}
