namespace Word2007RedactionTool
{
    partial class FormFindAndMark
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormFindAndMark));
            this.buttonCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox = new System.Windows.Forms.TextBox();
            this.buttonMark = new System.Windows.Forms.Button();
            this.panelOptions = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxIgnoreWhitespace = new System.Windows.Forms.CheckBox();
            this.checkBoxIgnorePunct = new System.Windows.Forms.CheckBox();
            this.checkBoxMatchSuffix = new System.Windows.Forms.CheckBox();
            this.checkBoxMatchPrefix = new System.Windows.Forms.CheckBox();
            this.checkBoxWordForms = new System.Windows.Forms.CheckBox();
            this.checkBoxSoundsLike = new System.Windows.Forms.CheckBox();
            this.checkBoxWildcards = new System.Windows.Forms.CheckBox();
            this.checkBoxWholeWord = new System.Windows.Forms.CheckBox();
            this.checkBoxMatchCase = new System.Windows.Forms.CheckBox();
            this.buttonMoreLess = new System.Windows.Forms.Button();
            this.labelResults = new System.Windows.Forms.Label();
            this.labelOptions = new System.Windows.Forms.Label();
            this.labelOptionsDetails = new System.Windows.Forms.Label();
            this.panelOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonCancel
            // 
            resources.ApplyResources(this.buttonCancel, "buttonCancel");
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // textBox
            // 
            resources.ApplyResources(this.textBox, "textBox");
            this.textBox.Name = "textBox";
            this.textBox.TextChanged += new System.EventHandler(this.textBox_TextChanged);
            // 
            // buttonMark
            // 
            resources.ApplyResources(this.buttonMark, "buttonMark");
            this.buttonMark.Name = "buttonMark";
            this.buttonMark.UseVisualStyleBackColor = true;
            this.buttonMark.Click += new System.EventHandler(this.buttonMark_Click);
            // 
            // panelOptions
            // 
            this.panelOptions.Controls.Add(this.label2);
            this.panelOptions.Controls.Add(this.checkBoxIgnoreWhitespace);
            this.panelOptions.Controls.Add(this.checkBoxIgnorePunct);
            this.panelOptions.Controls.Add(this.checkBoxMatchSuffix);
            this.panelOptions.Controls.Add(this.checkBoxMatchPrefix);
            this.panelOptions.Controls.Add(this.checkBoxWordForms);
            this.panelOptions.Controls.Add(this.checkBoxSoundsLike);
            this.panelOptions.Controls.Add(this.checkBoxWildcards);
            this.panelOptions.Controls.Add(this.checkBoxWholeWord);
            this.panelOptions.Controls.Add(this.checkBoxMatchCase);
            resources.ApplyResources(this.panelOptions, "panelOptions");
            this.panelOptions.Name = "panelOptions";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // checkBoxIgnoreWhitespace
            // 
            resources.ApplyResources(this.checkBoxIgnoreWhitespace, "checkBoxIgnoreWhitespace");
            this.checkBoxIgnoreWhitespace.Name = "checkBoxIgnoreWhitespace";
            this.checkBoxIgnoreWhitespace.UseVisualStyleBackColor = true;
            this.checkBoxIgnoreWhitespace.CheckedChanged += new System.EventHandler(this.checkBoxGeneral_CheckedChanged);
            // 
            // checkBoxIgnorePunct
            // 
            resources.ApplyResources(this.checkBoxIgnorePunct, "checkBoxIgnorePunct");
            this.checkBoxIgnorePunct.Name = "checkBoxIgnorePunct";
            this.checkBoxIgnorePunct.UseVisualStyleBackColor = true;
            this.checkBoxIgnorePunct.CheckedChanged += new System.EventHandler(this.checkBoxGeneral_CheckedChanged);
            // 
            // checkBoxMatchSuffix
            // 
            resources.ApplyResources(this.checkBoxMatchSuffix, "checkBoxMatchSuffix");
            this.checkBoxMatchSuffix.Name = "checkBoxMatchSuffix";
            this.checkBoxMatchSuffix.UseVisualStyleBackColor = true;
            this.checkBoxMatchSuffix.CheckedChanged += new System.EventHandler(this.checkBoxGeneral_CheckedChanged);
            // 
            // checkBoxMatchPrefix
            // 
            resources.ApplyResources(this.checkBoxMatchPrefix, "checkBoxMatchPrefix");
            this.checkBoxMatchPrefix.Name = "checkBoxMatchPrefix";
            this.checkBoxMatchPrefix.UseVisualStyleBackColor = true;
            this.checkBoxMatchPrefix.CheckedChanged += new System.EventHandler(this.checkBoxGeneral_CheckedChanged);
            // 
            // checkBoxWordForms
            // 
            resources.ApplyResources(this.checkBoxWordForms, "checkBoxWordForms");
            this.checkBoxWordForms.Name = "checkBoxWordForms";
            this.checkBoxWordForms.UseVisualStyleBackColor = true;
            this.checkBoxWordForms.CheckedChanged += new System.EventHandler(this.checkBoxFuzzy_CheckedChanged);
            // 
            // checkBoxSoundsLike
            // 
            resources.ApplyResources(this.checkBoxSoundsLike, "checkBoxSoundsLike");
            this.checkBoxSoundsLike.Name = "checkBoxSoundsLike";
            this.checkBoxSoundsLike.UseVisualStyleBackColor = true;
            this.checkBoxSoundsLike.CheckedChanged += new System.EventHandler(this.checkBoxFuzzy_CheckedChanged);
            // 
            // checkBoxWildcards
            // 
            resources.ApplyResources(this.checkBoxWildcards, "checkBoxWildcards");
            this.checkBoxWildcards.Name = "checkBoxWildcards";
            this.checkBoxWildcards.UseVisualStyleBackColor = true;
            this.checkBoxWildcards.CheckedChanged += new System.EventHandler(this.checkBoxFuzzy_CheckedChanged);
            // 
            // checkBoxWholeWord
            // 
            resources.ApplyResources(this.checkBoxWholeWord, "checkBoxWholeWord");
            this.checkBoxWholeWord.Name = "checkBoxWholeWord";
            this.checkBoxWholeWord.UseVisualStyleBackColor = true;
            this.checkBoxWholeWord.CheckedChanged += new System.EventHandler(this.checkBoxGeneral_CheckedChanged);
            // 
            // checkBoxMatchCase
            // 
            resources.ApplyResources(this.checkBoxMatchCase, "checkBoxMatchCase");
            this.checkBoxMatchCase.Name = "checkBoxMatchCase";
            this.checkBoxMatchCase.UseVisualStyleBackColor = true;
            this.checkBoxMatchCase.CheckedChanged += new System.EventHandler(this.checkBoxGeneral_CheckedChanged);
            // 
            // buttonMoreLess
            // 
            resources.ApplyResources(this.buttonMoreLess, "buttonMoreLess");
            this.buttonMoreLess.Name = "buttonMoreLess";
            this.buttonMoreLess.UseVisualStyleBackColor = true;
            this.buttonMoreLess.Click += new System.EventHandler(this.buttonMoreLess_Click);
            // 
            // labelResults
            // 
            resources.ApplyResources(this.labelResults, "labelResults");
            this.labelResults.Name = "labelResults";
            // 
            // labelOptions
            // 
            resources.ApplyResources(this.labelOptions, "labelOptions");
            this.labelOptions.Name = "labelOptions";
            // 
            // labelOptionsDetails
            // 
            this.labelOptionsDetails.AutoEllipsis = true;
            resources.ApplyResources(this.labelOptionsDetails, "labelOptionsDetails");
            this.labelOptionsDetails.Name = "labelOptionsDetails";
            // 
            // FormFindAndMark
            // 
            this.AcceptButton = this.buttonMark;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.Controls.Add(this.labelOptionsDetails);
            this.Controls.Add(this.labelOptions);
            this.Controls.Add(this.labelResults);
            this.Controls.Add(this.buttonMoreLess);
            this.Controls.Add(this.panelOptions);
            this.Controls.Add(this.buttonMark);
            this.Controls.Add(this.textBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormFindAndMark";
            this.ShowInTaskbar = false;
            this.Load += new System.EventHandler(this.FormFindAndMark_Load);
            this.panelOptions.ResumeLayout(false);
            this.panelOptions.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.Button buttonMark;
        private System.Windows.Forms.Panel panelOptions;
        private System.Windows.Forms.CheckBox checkBoxIgnoreWhitespace;
        private System.Windows.Forms.CheckBox checkBoxIgnorePunct;
        private System.Windows.Forms.CheckBox checkBoxMatchSuffix;
        private System.Windows.Forms.CheckBox checkBoxMatchPrefix;
        private System.Windows.Forms.CheckBox checkBoxWordForms;
        private System.Windows.Forms.CheckBox checkBoxSoundsLike;
        private System.Windows.Forms.CheckBox checkBoxWildcards;
        private System.Windows.Forms.CheckBox checkBoxWholeWord;
        private System.Windows.Forms.CheckBox checkBoxMatchCase;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonMoreLess;
        private System.Windows.Forms.Label labelResults;
        private System.Windows.Forms.Label labelOptions;
        private System.Windows.Forms.Label labelOptionsDetails;
    }
}