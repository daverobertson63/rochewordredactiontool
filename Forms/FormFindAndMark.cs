//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Word2007RedactionTool.Properties;

namespace Word2007RedactionTool
{
    public partial class FormFindAndMark : Form
    {
        Word.WdColor ShadingColor;

        public FormFindAndMark(Word.WdColor color)
        {
            ShadingColor = color;
            InitializeComponent();
        }

        private void FormFindAndMark_Load(object sender, EventArgs e)
        {
            //clear labels
            labelResults.Text = string.Empty;
            labelOptionsDetails.Text = string.Empty;

            //set focus
            textBox.Select(0, 0);

            //load up defaults
            checkBoxIgnorePunct.Checked = Settings.Default.FindIgnorePunct;
            checkBoxIgnoreWhitespace.Checked = Settings.Default.FindIgnoreSpace;
            checkBoxMatchCase.Checked = Settings.Default.FindMatchCase;
            checkBoxMatchPrefix.Checked = Settings.Default.FindMatchPrefix;
            checkBoxMatchSuffix.Checked = Settings.Default.FindMatchSuffix;
            checkBoxSoundsLike.Checked = Settings.Default.FindSoundsLike;
            checkBoxWholeWord.Checked = Settings.Default.FindWholeWords;
            checkBoxWildcards.Checked = Settings.Default.FindWildcards;
            checkBoxWordForms.Checked = Settings.Default.FindAllWordForms;            
            if (Settings.Default.FindDetails)
            {
                panelOptions.Visible = true;
                this.Height += panelOptions.Height;
                buttonMoreLess.Text = Resources.ButtonLess;
            }

            UpdateOptionsStrings();
        }

        private void buttonMoreLess_Click(object sender, EventArgs e)
        {
            if(panelOptions.Visible == false)
            {
                panelOptions.Visible = true;
                this.Height += panelOptions.Height;
                buttonMoreLess.Text = Resources.ButtonLess;
                Settings.Default.FindDetails = true;
            }
            else
            {
                panelOptions.Visible = false;
                this.Height -= panelOptions.Height;
                buttonMoreLess.Text = Resources.ButtonMore;
                Settings.Default.FindDetails = false;
            }
        }

        private void buttonMark_Click(object sender, EventArgs e)
        {
            int HitCount = 0;
            try
            {
                //search
                foreach (Word.Range StoryRange in Globals.ThisAddIn.Application.Selection.Document.StoryRanges)
                {
                    HitCount += FindAndMarkInStory(StoryRange);
                }

                //show results
                if (HitCount > 1) 
                {
                    labelResults.Text = HitCount + Resources.OccurrencesFound;
                }
                else if (HitCount == 1)
                {
                    labelResults.Text = HitCount + Resources.OccurrenceFound;
                }
                else //we didn't find anything, tell them.
                {
                    labelResults.Text = string.Empty;
                    GenericMessageBox.Show(Resources.TextNotFound, Resources.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
                }
            }
            catch (COMException cex)
            {
                //something failed, show them the error message
                GenericMessageBox.Show(cex.Message, Resources.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            //save settings
            Settings.Default.FindIgnorePunct = checkBoxIgnorePunct.Checked;
            Settings.Default.FindIgnoreSpace = checkBoxIgnoreWhitespace.Checked;
            Settings.Default.FindMatchCase = checkBoxMatchCase.Checked;
            Settings.Default.FindMatchPrefix = checkBoxMatchPrefix.Checked;
            Settings.Default.FindMatchSuffix = checkBoxMatchSuffix.Checked;
            Settings.Default.FindSoundsLike = checkBoxSoundsLike.Checked;
            Settings.Default.FindWholeWords = checkBoxWholeWord.Checked;
            Settings.Default.FindWildcards = checkBoxWildcards.Checked;
            Settings.Default.FindAllWordForms = checkBoxWordForms.Checked;  
            Settings.Default.Save();

            this.Close();
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            //clear the results
            labelResults.Text = string.Empty;

            //check if the mark button should be enabled
            if (string.IsNullOrEmpty(textBox.Text))
                buttonMark.Enabled = false;
            else
                buttonMark.Enabled = true;

            //check if "whole word" should be enabled
            if (textBox.Text.Contains(" "))
            {
                checkBoxWholeWord.Checked = false;
                checkBoxWholeWord.Enabled = false;
            }
            else if(checkBoxMatchCase.Enabled) // Match Case is disabled ONLY in cases when this control cannot be enabled
            {
                checkBoxWholeWord.Enabled = true;
            }
        }

        private void checkBoxGeneral_CheckedChanged(object sender, EventArgs e)
        {
            UpdateOptionsStrings();
        }

        private void checkBoxFuzzy_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox CurrentCheckBox = (CheckBox)sender;

            //uncheck the disabled checkboxes
            checkBoxMatchPrefix.Checked = false;
            checkBoxMatchSuffix.Checked = false;
            checkBoxMatchCase.Checked = false;
            checkBoxWholeWord.Checked = false;

            //check if the disabled checkboxes should be toggled, if so toggle them
            if (CurrentCheckBox.Checked || !checkBoxWordForms.Checked && !checkBoxWildcards.Checked && !checkBoxSoundsLike.Checked)
            {
                checkBoxMatchPrefix.Enabled = !CurrentCheckBox.Checked;
                checkBoxMatchSuffix.Enabled = !CurrentCheckBox.Checked;
                checkBoxMatchCase.Enabled = !CurrentCheckBox.Checked;
                if (!textBox.Text.Contains(" "))
                    checkBoxWholeWord.Enabled = !CurrentCheckBox.Checked;
            }

            //set the state of the mutually exclusive settings
            if (CurrentCheckBox.Checked)
            {
                switch (CurrentCheckBox.Name)
                {
                    case "checkBoxWordForms":
                        checkBoxSoundsLike.Checked = false;
                        checkBoxWildcards.Checked = false;
                        break;
                    case "checkBoxWildcards":
                        checkBoxSoundsLike.Checked = false;
                        checkBoxWordForms.Checked = false;
                        break;
                    case "checkBoxSoundsLike":
                        checkBoxWildcards.Checked = false;
                        checkBoxWordForms.Checked = false;
                        break;
                    default:
                        System.Diagnostics.Debug.Fail("unknown check box");
                        break;
                }
            }

            UpdateOptionsStrings();
        }

        /// <summary>
        /// Find and mark all instances of a string in a specific StoryRange.
        /// </summary>
        /// <param name="StoryRange">A Range specifying the story range to search.</param>
        /// <returns>True if a hit was found, False otherwise.</returns>
        private int FindAndMarkInStory(Word.Range StoryRange)
        {
            object CollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            RangeData LastHit = new RangeData();
            int HitCount = 0;
            object Missing = Type.Missing;
            object FindText = textBox.Text;

            //text boxes in headers/footers aren't in the text box story
            if ((int)StoryRange.StoryType > 5 && StoryRange.ShapeRange.Count > 0)
            {
                foreach (Word.Shape Shape in StoryRange.ShapeRange)
                {
                    Word.Range ShapeRange = RedactCommon.RangeFromShape(Shape);
                    if (ShapeRange != null)
                        HitCount += FindAndMarkInStory(ShapeRange);
                }
            }

            Word.Find FindScope = StoryRange.Find;

            //set search options
            FindScope.IgnorePunct = checkBoxIgnorePunct.Checked;
            FindScope.IgnoreSpace = checkBoxIgnoreWhitespace.Checked;
            FindScope.MatchAllWordForms = checkBoxWordForms.Checked;
            FindScope.MatchCase = checkBoxMatchCase.Checked;
            FindScope.MatchPrefix = checkBoxMatchPrefix.Checked;
            FindScope.MatchSoundsLike = checkBoxSoundsLike.Checked;
            FindScope.MatchSuffix = checkBoxMatchSuffix.Checked;
            FindScope.MatchWholeWord = checkBoxWholeWord.Checked;
            FindScope.MatchWildcards = checkBoxWildcards.Checked;

            while (FindScope.Execute(ref FindText, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing) && (StoryRange.Start != LastHit.Start || StoryRange.End != LastHit.End))
            {
                HitCount++;
                LastHit = new RangeData(StoryRange.Start, StoryRange.End);
                StoryRange.Font.Shading.BackgroundPatternColor = (Word.WdColor)ShadingColor;
                StoryRange.Collapse(ref CollapseEnd);
            }
 
            //check in any subsequent stories
            if (StoryRange.NextStoryRange != null)
                HitCount += FindAndMarkInStory(StoryRange.NextStoryRange);

            return HitCount;
        }

        /// <summary>
        /// Update the Options text in the Find and Mark dialog box.
        /// </summary>
        private void UpdateOptionsStrings()
        {
            labelOptionsDetails.Text = string.Empty;

            if (checkBoxMatchCase.Checked)
            {
                labelOptionsDetails.Text += checkBoxMatchCase.Text;
            }
            if (checkBoxWholeWord.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxWholeWord.Text;
            }
            if (checkBoxWildcards.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxWildcards.Text;
            }
            if (checkBoxSoundsLike.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxSoundsLike.Text;
            }
            if (checkBoxWordForms.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxWordForms.Text;
            }
            if (checkBoxMatchPrefix.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxMatchPrefix.Text;
            }
            if (checkBoxMatchSuffix.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxMatchSuffix.Text;
            }
            if (checkBoxIgnorePunct.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxIgnorePunct.Text;
            }
            if (checkBoxIgnoreWhitespace.Checked)
            {
                if (!string.IsNullOrEmpty(labelOptionsDetails.Text))
                    labelOptionsDetails.Text += ", ";
                labelOptionsDetails.Text += checkBoxIgnoreWhitespace.Text;
            }

            //don't show the options if the details are blank
            if (string.IsNullOrEmpty(labelOptionsDetails.Text))
                labelOptions.Visible = false;
            else
                labelOptions.Visible = true;
        }        
    }
}
