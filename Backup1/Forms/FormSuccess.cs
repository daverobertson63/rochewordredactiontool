// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Word2007RedactionTool
{
    public partial class FormSuccess : Form
    {
        Word.Application Application;

        public FormSuccess(Word.Application wordApp)
        {
            InitializeComponent();

            Application = wordApp;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonInspect_Click(object sender, EventArgs e)
        {
            this.Close();

            object Missing = Type.Missing;
            Application.Dialogs[Word.WdWordDialog.wdDialogDocumentInspector].Show(ref Missing);
        }
    }
}
