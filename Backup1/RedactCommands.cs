// Copyright (c) Microsoft Corporation.  All rights reserved.
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Word2007RedactionTool.Properties;
using System;

namespace Word2007RedactionTool
{
    public partial class RedactRibbon
    {
        #region Ribbon Commands

        /// <summary>
        /// Mark the current selection.
        /// </summary>
        private void MarkSelection()
        {
            if (Application.Selection.Type == Word.WdSelectionType.wdSelectionIP)
                Application.Selection.Words[1].Font.Shading.BackgroundPatternColor = (Word.WdColor)ShadingColor;
            else
            {
                Application.Selection.Font.Shading.BackgroundPatternColor = (Word.WdColor)ShadingColor;
            }
        }

        /// <summary>
        /// Unmarks the current selection.
        /// </summary>
        private void UnmarkSelection()
        {
            Word.Range OriginalSelection = Application.Selection.Range.Duplicate;

            if (Application.Selection.Type == Word.WdSelectionType.wdSelectionIP)
                UnmarkRange(Application.Selection.Words[1], false);
            else
                UnmarkRange(Application.Selection.Range, false);

            OriginalSelection.Select();
        }

        /// <summary>
        /// Unmark the entire document.
        /// </summary>
        private void UnmarkDocument()
        {
            Application.ScreenUpdating = false;

            Word.Range OriginalSelection = Application.Selection.Range.Duplicate;
            Word.WdViewType OriginalView = Application.ActiveWindow.View.Type;

            foreach (Word.Range StoryRange in Application.Selection.Document.StoryRanges)
            {
                UnmarkRange(StoryRange, true);
            }

            Application.ActiveWindow.View.Type = OriginalView;
            OriginalSelection.Select();

            Application.ScreenUpdating = true;
        }

        /// <summary>
        /// Selects the previous mark (relative to the current selection).
        /// </summary>
        private void SelectPreviousMark()
        {
            Application.ScreenUpdating = false;
            Word.WdViewType OriginalView = Application.ActiveWindow.View.Type;

            Word.Selection CurrentSelection = Application.Selection;
            Word.Range OriginalSelection = CurrentSelection.Range.Duplicate;

            //collapse to the start of the current selection
            Word.Range SelectionRange = CurrentSelection.Range;
            SelectionRange.Collapse(ref CollapseStart);
            ExtendPreviousMark(ref SelectionRange, true);
            
            if (!FindPreviousMark(ref SelectionRange, OriginalView, (SelectionRange.StoryType != Word.WdStoryType.wdEndnoteSeparatorStory || SelectionRange.End < SelectionRange.Document.StoryRanges[Word.WdStoryType.wdEndnoteSeparatorStory].StoryLength-1 )))
                OriginalSelection.Select();

            Application.ScreenUpdating = true;
        }

        /// <summary>
        /// Selects the next mark (relative to the current selection).
        /// </summary>
        private void SelectNextMark()
        {
            Application.ScreenUpdating = false;
            Word.WdViewType OriginalView = Application.ActiveWindow.View.Type;

            Word.Selection CurrentSelection = Application.Selection;
            Word.Range OriginalSelection = CurrentSelection.Range.Duplicate;            

            //collapse to the end of the current selection
            Word.Range SelectionRange = CurrentSelection.Range;
            SelectionRange.Collapse(ref CollapseEnd);
            ExtendNextMark(ref SelectionRange, true);
            
            if (!FindNextMark(ref SelectionRange, OriginalView, (CurrentSelection.StoryType != Word.WdStoryType.wdMainTextStory || CurrentSelection.End != 0)/* don't wrap from cp0 */))
                OriginalSelection.Select();

            Application.ScreenUpdating = true;
        }

        /// <summary>
        /// Redact the current document.
        /// </summary>
        private void RedactDocument()
        {
            object Story = Word.WdUnits.wdStory;
            Word.WdViewType OriginalView = Application.ActiveWindow.View.Type;
            Application.ScreenUpdating = false;

            Word.Document SourceFile = Application.ActiveDocument;

            //cache the document's saved and update styles from template states
            bool Saved = SourceFile.Saved;
            bool UpdateStyles = SourceFile.UpdateStylesOnOpen;

            //BUG 5819: need to make sure this is off so that Normal's properties don't get demoted into redacted file
            Application.ActiveDocument.UpdateStylesOnOpen = false; 
            Word.Document FileToRedact = RedactCommon.CloneDocument(Application.ActiveDocument);

            //reset those states to their cached values
            SourceFile.UpdateStylesOnOpen = UpdateStyles;
            SourceFile.Saved = Saved; 

            //get the window for the document surface (_WwG)
            NativeWindow WordWindow = RedactCommon.GetDocumentSurfaceWindow();

            //show progress UI
            using (FormDoRedaction ProgressUI = new FormDoRedaction(FileToRedact, this))
            {
                ProgressUI.ResetProgress();
                DialogResult RedactResult = ProgressUI.ShowDialog();

                //fix the view
                Application.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
                Application.ActiveWindow.View.Type = OriginalView;
                Application.Selection.HomeKey(ref Story, ref Missing);

                if (RedactResult != DialogResult.OK)
                {
                    GenericMessageBox.Show(Resources.RedactionFailed, Resources.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
                }
                else
                {
                    //dialog telling the user we're done
                    using (FormSuccess Success = new FormSuccess(FileToRedact.Application))
                        Success.ShowDialog(WordWindow);
                }
            }

            //set focus back onto the document
            NativeMethods.SetFocus(WordWindow.Handle);

            //we need to release our handle, or Word will crash on close as the CLR tries to do it for us
            WordWindow.ReleaseHandle();
        }

        /// <summary>
        /// Show the Find and Mark dialog box.
        /// </summary>
        private void FindAndMark()
        {
            Word.Range OriginalSelection = Application.Selection.Range.Duplicate;
            Word.WdViewType OriginalView = Application.ActiveWindow.View.Type;

            //get the window for the document surface (_WwG)
            NativeWindow WordWindow = RedactCommon.GetDocumentSurfaceWindow();

            using (FormFindAndMark find = new FormFindAndMark(ShadingColor))
            {
                find.ShowDialog(WordWindow);
            }

            //set focus back onto the document
            NativeMethods.SetFocus(WordWindow.Handle);

            //we need to release our handle, or Word will crash on close as the CLR tries to do it for us
            WordWindow.ReleaseHandle();

            Application.ActiveWindow.View.Type = OriginalView;
            OriginalSelection.Select();
        }

        #endregion

        /// <summary>
        /// Removes redaction marks from the specified range.
        /// </summary>
        /// <param name="Range">The Range from which to clear redaction marks.</param>
        /// <param name="UnmarkAll">True to also clear any subranges (e.g. text boxes), False otherwise.</param>
        private void UnmarkRange(Word.Range Range, bool UnmarkAll)
        {
            //text boxes in headers/footers aren't in the text box story
            if (UnmarkAll && (int)Range.StoryType > 5 && Range.ShapeRange.Count > 0)
            {
                foreach (Word.Shape Shape in Range.ShapeRange)
                {
                    Word.Range ShapeRange = RedactCommon.RangeFromShape(Shape);
                    if (ShapeRange != null)
                        UnmarkRange(ShapeRange, true);
                }
            }

            //unmark the range
            List<RangeDataEx> RangeMarkers = RedactCommon.GetAllMarkedRanges(Range, ShadingColor, true);
            foreach (RangeDataEx UniqueRange in RangeMarkers)
            {
                Range.Start = UniqueRange.Start;
                Range.End = UniqueRange.End;

                Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
            }

            //also catch other instances of the same StoryRange
            if (UnmarkAll && Range.NextStoryRange != null)
                UnmarkRange(Range.NextStoryRange, true);
        }

        /// <summary>
        /// Redact the specified document.
        /// </summary>
        /// <param name="Document">The document to be redacted.</param>
        /// <returns>True if redaction succeeded, otherwise False.</returns>
        internal bool RedactDocument(Word.Document Document, BackgroundWorker Worker)
        {
            try
            {
                //redact each story
                foreach (Word.Range StoryRange in Document.StoryRanges)
                {
                    RedactStoryRange(StoryRange, Worker);
                }
            }
            catch (Exception) //moving to catch a generic exception in order to ensure the redaction thread never dies and leaves the progress meter "stuck"
            {
                return false;
            }
            return true;
        }
       
        /// <summary>
        /// Redact a story range (e.g. all textboxes, all first page headers, the main document).
        /// </summary>
        /// <param name="StoryRange">The story range to redact.</param>
        /// <param name="Worker">A BackgroundWorker on which to report progress.</param>
        private void RedactStoryRange(Word.Range StoryRange, BackgroundWorker Worker)
        {
            //textboxes in headers/footers/textboxes/etc. are not in the textbox story.
            if ((int)StoryRange.StoryType > 5 && StoryRange.ShapeRange.Count > 0)
            {
                foreach (Word.Shape Shape in StoryRange.ShapeRange)
                {
                    Word.Range ShapeRange = RedactCommon.RangeFromShape(Shape);
                    if (ShapeRange != null)
                        RedactStoryRange(ShapeRange, Worker);
                }
            }

            //redact all shapes
            RedactShapes(StoryRange);

            //remove redaction marks from all paragraph mark glyphs
            Word.Range ParaMark = StoryRange.Duplicate;
            foreach (Word.Paragraph Paragraph in StoryRange.Paragraphs)
            {
                //select the para mark
                ParaMark.Start = Paragraph.Range.End - 1;
                ParaMark.End = ParaMark.Start + 1;

                if (RedactCommon.IsMarkedRange(ParaMark, ShadingColor))
                    ParaMark.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
            }

            //run through the collection            
            //since we're messing with the text, we need to go back to front.
            StoryRange.Collapse(ref CollapseEnd);
            while (FindPreviousMarkInCurrentStoryRange(ref StoryRange))
            {
                Debug.Assert(StoryRange.Paragraphs.Count < 2, "redacting more than one paragraph - Selection.Move will behave incorrectly in tables");

                //break up the range by formatting, so we can maintain layout through the redaction
                List<RangeDataEx> RangeMarkers = RedactCommon.GetAllMarkedRanges(StoryRange, ShadingColor, false);
                for (int i = RangeMarkers.Count - 1; i >= 0; i--)
                {
                    StoryRange.Start = RangeMarkers[i].Start;
                    StoryRange.End = RangeMarkers[i].End;

                    RedactRange(StoryRange, RangeMarkers[i]);

                    //update progress UI
                    if (Worker != null && StoryRange.StoryType == Word.WdStoryType.wdMainTextStory)
                        Worker.ReportProgress((((StoryRange.StoryLength - RangeMarkers[i].Start) * 100) / StoryRange.StoryLength), null);
                }
            }

            //get all other stories
            if (StoryRange.NextStoryRange != null)
                RedactStoryRange(StoryRange.NextStoryRange, Worker);

            //give each non-main story a small % of the progress bar
            if (Worker != null && StoryRange.StoryType != Word.WdStoryType.wdMainTextStory)
                Worker.ReportProgress(100 + (int)StoryRange.StoryType);
        }

        /// <summary>
        /// Redact all shapes in a range.
        /// </summary>
        /// <param name="StoryRange">A range containing zero or more shapes.</param>
        private void RedactShapes(Word.Range StoryRange)
        {
            if (StoryRange.ShapeRange.Count > 0)
            {
                List<int> ShapeLocations = new List<int>();
                Dictionary<int, Word.Shape> Shapes = new Dictionary<int, Word.Shape>();

                //scan for shapes
                //since we're messing with the text, we need reorder them back to front by document order (they're in z-order).
                foreach (Word.Shape Shape in StoryRange.ShapeRange)
                {
                    Word.Range AnchorRange = Shape.Anchor;
                    ShapeLocations.Add(AnchorRange.Start);
                    Shapes.Add(AnchorRange.Start, Shape);
                }

                //sort the anchor locations
                ShapeLocations.Sort();

                for (int i = ShapeLocations.Count - 1; i >= 0; i--)
                {
                    Word.Shape Shape = Shapes[ShapeLocations[i]];
                    if (RedactCommon.IsMarkedRange(Shape.Anchor, ShadingColor))
                    {
                        Debug.WriteLine("Shape from " + Shape.Anchor.Start + " to " + Shape.Anchor.End + "to be redacted.");
                        RedactShape(StoryRange.Document, Shape);
                    }
                }
            }
        }

        /// <summary>
        /// Redact a range, replacing all marked text.
        /// </summary>
        /// <param name="range">A range to redact.</param>
        /// <param name="rangeData">A RangeDataEx containing properties about the range.</param>
        private void RedactRange(Word.Range range, RangeDataEx rangeData)
        {
            object ParagraphNumber = Word.WdNumberType.wdNumberParagraph;

            foreach (Word.Paragraph p in range.Paragraphs)
            {
                //get the para's range
                Word.Range ParagraphRange = p.Range;

                //trim to the selection, if needed
                if (range.Start > ParagraphRange.Start)
                    ParagraphRange.Start = range.Start;
                if (range.End < ParagraphRange.End)
                    ParagraphRange.End = range.End;
                if (ParagraphRange.End == p.Range.End - 1 && ParagraphRange.Start == p.Range.Start)
                    p.Range.ListFormat.ConvertNumbersToText(ref ParagraphNumber); //if the whole para was redacted, redact the numbering

                //make it black on black
                ParagraphRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                ParagraphRange.HighlightColorIndex = Word.WdColorIndex.wdBlack; //moving to highlighting instead of text background
                ParagraphRange.Font.Color = Word.WdColor.wdColorBlack;

                //get rid of links and bookmarks
                foreach (Word.Hyperlink Hyperlink in ParagraphRange.Hyperlinks)
                    Hyperlink.Delete();
                foreach (Word.Bookmark Bookmark in ParagraphRange.Bookmarks)
                    Bookmark.Delete();

                //BUG 110: suppress proofing errors
                ParagraphRange.NoProofing = -1;

                //finally, replace the text
                Debug.Assert(ParagraphRange.ShapeRange.Count == 0, "Some Shapes were not redacted by RedactShapes.");
                if (ParagraphRange.InlineShapes.Count > 0)
                {
                    //if there are images, then split into subranges and process text and images separately
                    List<RangeData> Subranges = RedactCommon.SplitRange(ParagraphRange);

                    for (int j = Subranges.Count - 1; j >= 0; j--)
                    {
                        //set start and end
                        ParagraphRange.Start = Subranges[j].Start;
                        ParagraphRange.End = Subranges[j].End;

                        if (ParagraphRange.InlineShapes.Count > 0)
                            RedactInlineShape(ParagraphRange.InlineShapes[1]);
                        else
                            ParagraphRange.Text = RedactCommon.BuildFillerText(ParagraphRange, rangeData);
                    }
                }
                else
                    ParagraphRange.Text = RedactCommon.BuildFillerText(ParagraphRange, rangeData);
            }
        }

        /// <summary>
        /// Redact an inline shape by replacing it with a 1px image.
        /// </summary>
        /// <param name="InlineShape">The InlineShape to redact.</param>
        private void RedactInlineShape(Word.InlineShape InlineShape)
        {
            string ImagePath = RedactCommon.CreateRedactedImage();

            float Height = InlineShape.Height;
            float Width = InlineShape.Width;

            InlineShape.Range.Select();
            Application.Selection.Delete(ref Missing, ref Missing);

            Word.InlineShape RedactedImage = Application.Selection.InlineShapes.AddPicture(ImagePath, ref Missing, ref Missing, ref Missing);
            RedactedImage.AlternativeText = string.Empty;
            RedactedImage.Height = Height;
            RedactedImage.Width = Width;
        }       

        /// <summary>
        /// Redact a shapes by replacing it with a 1px image.
        /// </summary>
        /// <param name="Document">The document containing the shape.</param>
        /// <param name="ShapeToRedact">The shape to redact.</param>
        private void RedactShape(Word.Document Document, Word.Shape ShapeToRedact)
        {
            string ImagePath = RedactCommon.CreateRedactedImage();
            object AnchorRange = ShapeToRedact.Anchor;
            object Height = ShapeToRedact.Height;
            object Width = ShapeToRedact.Width;
            object Top = ShapeToRedact.Top;
            object Left = ShapeToRedact.Left;
            object WrapType = ShapeToRedact.WrapFormat.Type;

            ((Word.Range)AnchorRange).Select();
            ShapeToRedact.Delete();

            Word.Shape RedactedImage = Document.Shapes.AddPicture(ImagePath, ref Missing, ref Missing, ref Left, ref Top, ref Width, ref Height, ref AnchorRange);
            RedactedImage.AlternativeText = string.Empty;
            RedactedImage.WrapFormat.Type = (Word.WdWrapType)WrapType;
        }
    }       
}
