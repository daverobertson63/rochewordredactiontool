// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Word2007RedactionTool.Properties;

namespace Word2007RedactionTool
{
    partial class RedactRibbon
    {
        /// <summary>
        /// Finds the previous redaction mark in the document.
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <param name="CurrentView">The current view type.</param>
        /// <param name="WrapAtStart">True to prompt to wrap around at the start, otherwise False.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private bool FindPreviousMark(ref Word.Range CurrentRange, Word.WdViewType CurrentView, bool WrapAtStart)
        {
            if (FindPreviousMarkInCurrentStory(ref CurrentRange) || FindPreviousMarkInOtherStory(ref CurrentRange))
            {
                //extend it out and select it
                ExtendPreviousMark(ref CurrentRange, false);
                CurrentRange.Select();
                return true;
            }
            else
            {
                if (WrapAtStart)
                {
                    if (GenericMessageBox.Show(Resources.SearchFromEnd, Resources.AppName, MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0) == DialogResult.OK)
                    {
                        //they want to wrap around, so find the last story and start again
                        CurrentRange = GetLastStory(ref CurrentRange);
                        CurrentRange.Collapse(ref CollapseEnd);
                        
                        return FindPreviousMark(ref CurrentRange, CurrentView, false);
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    GenericMessageBox.Show(Resources.NoRedactionMarks, Resources.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
                    return false;
                }
            }
        }

        /// <summary>
        /// Finds the previous redaction mark in current document story (e.g. this or any previous odd page footer).
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private static bool FindPreviousMarkInCurrentStory(ref Word.Range CurrentRange)
        {
            //execute the search
            if (FindPreviousMarkInCurrentStoryRange(ref CurrentRange))
                return true;
            else if (RedactCommon.PreviousStoryRange(CurrentRange) != null)
                return FindPreviousMarkInStoryRanges(ref CurrentRange);
            else
                return false;
        }

        /// <summary>
        /// Finds the previous redaction mark in current document story range (e.g. the current text box).
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private static bool FindPreviousMarkInCurrentStoryRange(ref Word.Range CurrentRange)
        {
            object Missing = Type.Missing;
            object False = false;
            Word.Find FindScope = CurrentRange.Find;
            FindScope.Font.Shading.BackgroundPatternColor = ShadingColor;

            return FindScope.ExecuteOld(ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref False, ref Missing, ref Missing, ref Missing, ref Missing);
        }

        /// <summary>
        /// Finds the previous redaction mark in previous document stories of the current type (e.g. previous headers).
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private static bool FindPreviousMarkInStoryRanges(ref Word.Range CurrentRange)
        {
            object CollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;

            Word.WdHeaderFooterIndex? headerFooterId = null;

            Word.Range PreviousStoryRange = RedactCommon.PreviousStoryRange(CurrentRange);
            while (PreviousStoryRange != null)
            {
                //if we're in headers/footers, we need to get our location so we can update the current view
                if (CurrentRange.StoryType != Word.WdStoryType.wdTextFrameStory)
                    headerFooterId = RedactCommon.GetHeaderFooterType(CurrentRange.StoryType);

                CurrentRange = PreviousStoryRange;
                PreviousStoryRange = RedactCommon.PreviousStoryRange(CurrentRange);
                CurrentRange.Collapse(ref CollapseEnd);
                if (FindPreviousMarkInCurrentStoryRange(ref CurrentRange))
                {
                    //move the focus in the view
                    if (RedactCommon.IsHeaderFooter(CurrentRange.StoryType))
                        MoveHeaderFooterFocus(CurrentRange, RedactCommon.FindSectionWithHeaderFooter(CurrentRange.Document, CurrentRange), headerFooterId);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Finds the previous redaction mark in previous document stories.
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private bool FindPreviousMarkInOtherStory(ref Word.Range CurrentRange)
        {
            //move backward a story
            int PrevStory = (int)CurrentRange.StoryType - 1;

            while (PrevStory >= 1)
            {
                try
                {
                    CurrentRange = CurrentRange.Document.StoryRanges[(Word.WdStoryType)PrevStory];
                    while (CurrentRange.NextStoryRange != null)
                        CurrentRange = CurrentRange.NextStoryRange;
                    CurrentRange.Collapse(ref CollapseEnd);                   
                }
                catch (COMException)
                { }

                if (FindPreviousMarkInCurrentStory(ref CurrentRange))
                {
                    if (RedactCommon.IsHeaderFooter((Word.WdStoryType)PrevStory))
                        MoveHeaderFooterFocus(CurrentRange, RedactCommon.FindSectionWithHeaderFooter(CurrentRange.Document, CurrentRange), RedactCommon.GetHeaderFooterType((Word.WdStoryType)PrevStory));
                    else if (CurrentRange.Document.ActiveWindow.View.Type == Word.WdViewType.wdPrintView)
                        CurrentRange.Document.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument; //set the seek back to the document
                    return true;
                }

                PrevStory--;
            }

            return false;
        }

        /// <summary>
        /// Extends the current range's starting point to the beginning of the mark.
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the full mark if one is found.</param>
        /// <param name="Collapse">True to collapse the result to the start of the range, otherwise False.</param>
        private static void ExtendPreviousMark(ref Word.Range SelectionRange, bool Collapse)
        {
            if (SelectionRange.Start != 0)
            {
                //move before the current mark (if we are in one)
                if (RedactCommon.IsMarkedRange(SelectionRange, ShadingColor))
                {
                    Word.Range SearchRange = SelectionRange.Duplicate;
                    while (FindPreviousMarkInCurrentStoryRange(ref SearchRange))
                    {
                        if (SearchRange.End == SelectionRange.Start && (bool)SearchRange.get_Information(Word.WdInformation.wdWithInTable) == (bool)SelectionRange.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            SelectionRange.Start = SearchRange.Start;
                            if (Collapse)
                                SelectionRange.End = SearchRange.Start;
                        }
                        else
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Find the last accessible document story.
        /// </summary>
        /// <param name="CurrentRange">The starting point of the search.</param>
        /// <returns>The Range containing the last accessible story.</returns>
        private static Word.Range GetLastStory(ref Word.Range CurrentRange)
        {
            int StoryId = 17;
            while (StoryId >= 1)
            {
                try
                {
                    return CurrentRange.Document.StoryRanges[(Word.WdStoryType)StoryId];
                }
                catch (COMException)
                { }
                finally
                {
                    StoryId--;
                }
            }
            throw new InvalidOperationException("couldn't find a single valid story");
        }

        /// <summary>
        /// Finds the next redaction mark in the document.
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <param name="CurrentView">The current view type.</param>
        /// <param name="WrapAtEnd">True to prompt to wrap around at the end, otherwise False.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private bool FindNextMark(ref Word.Range CurrentRange, Word.WdViewType CurrentView, bool WrapAtEnd)
        {
            if (FindNextMarkInCurrentStory(ref CurrentRange) || FindNextMarkInOtherStory(ref CurrentRange))
            {
                //extend it out and select it
                ExtendNextMark(ref CurrentRange, false);
                CurrentRange.Select();
                return true;
            }
            else
            {
                if (WrapAtEnd)
                {
                    if (GenericMessageBox.Show(Resources.SearchFromBeginning, Resources.AppName, MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0) == DialogResult.OK)
                    {
                        //move the range to the start of the document and search again
                        CurrentRange = CurrentRange.Document.StoryRanges[Word.WdStoryType.wdMainTextStory];
                        CurrentRange.Collapse(ref CollapseStart);
                        return FindNextMark(ref CurrentRange, CurrentView, false);
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    GenericMessageBox.Show(Resources.NoRedactionMarks, Resources.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
                    return false;
                }
            }
        }

        /// <summary>
        /// Finds the next redaction mark in current document story (e.g. this or any following odd page footer).
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private static bool FindNextMarkInCurrentStory(ref Word.Range CurrentRange)
        {
            Word.Range OriginalRange = CurrentRange.Duplicate;

            //execute the search
            if (!(CurrentRange.End == CurrentRange.StoryLength || CurrentRange.End == CurrentRange.StoryLength - 1) // don't search this story if we're at the end
                && FindNextMarkInCurrentStoryRange(ref CurrentRange) && (OriginalRange.Start != CurrentRange.Start || OriginalRange.End != CurrentRange.End) /* did we move the selection in Find? */)
                return true;
            else if (CurrentRange.NextStoryRange != null)
                return FindNextMarkInStoryRanges(ref CurrentRange);
            else
                return false;
        }

        /// <summary>
        /// Finds the next redaction mark in current document story range (e.g. the current text box).
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private static bool FindNextMarkInCurrentStoryRange(ref Word.Range CurrentRange)
        {
            object Missing = Type.Missing;
            Word.Find FindScope = CurrentRange.Find;
            FindScope.Font.Shading.BackgroundPatternColor = ShadingColor;

            return FindScope.ExecuteOld(ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing, ref Missing);
        }

        /// <summary>
        /// Finds the next redaction mark in subsequent document stories of the current type (e.g. subsequent headers).
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private static bool FindNextMarkInStoryRanges(ref Word.Range SelectionRange)
        {
            object CollapseStart = Word.WdCollapseDirection.wdCollapseStart;

            Word.WdHeaderFooterIndex? headerFooterId = null;

            while (SelectionRange.NextStoryRange != null)
            {
                //get the header/footer ID
                if (SelectionRange.StoryType != Word.WdStoryType.wdTextFrameStory)
                    headerFooterId = RedactCommon.GetHeaderFooterType(SelectionRange.StoryType);

                SelectionRange = SelectionRange.NextStoryRange;
                SelectionRange.Collapse(ref CollapseStart);
                if (FindNextMarkInCurrentStoryRange(ref SelectionRange))
                {
                    //move the focus in the view
                    if (RedactCommon.IsHeaderFooter(SelectionRange.StoryType))
                        MoveHeaderFooterFocus(SelectionRange, RedactCommon.FindSectionWithHeaderFooter(SelectionRange.Document, SelectionRange), headerFooterId);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Finds the next redaction mark in subsequent document stories.
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the output mark if one is found.</param>
        /// <returns>True if a mark was found, otherwise False.</returns>
        private bool FindNextMarkInOtherStory(ref Word.Range CurrentRange)
        {
            //move forward a story
            int NextStory = (int)CurrentRange.StoryType + 1;

            while (NextStory <= 17)
            {
                try
                {
                    CurrentRange = CurrentRange.Document.StoryRanges[(Word.WdStoryType)NextStory];
                    CurrentRange.Collapse(ref CollapseStart);

                    if (FindNextMarkInCurrentStory(ref CurrentRange))
                    {
                        if (RedactCommon.IsHeaderFooter((Word.WdStoryType)NextStory))
                            MoveHeaderFooterFocus(CurrentRange, RedactCommon.FindSectionWithHeaderFooter(CurrentRange.Document, CurrentRange), RedactCommon.GetHeaderFooterType((Word.WdStoryType)NextStory));
                        return true;
                    }
                }
                catch (COMException)
                { }

                NextStory++;
            }

            return false;
        }

        /// <summary>
        /// Extends the current range's ending point to the end of the mark.
        /// </summary>
        /// <param name="CurrentRange">The range containing the starting point of the search. Changes to the location of the full mark if one is found.</param>
        /// <param name="Collapse">True to collapse the result to the end of the range, otherwise False.</param>
        private static void ExtendNextMark(ref Word.Range SelectionRange, bool Collapse)
        {
            if (SelectionRange.End < SelectionRange.StoryLength-1)
            {
                //move before the current mark (if we are in one)
                if (RedactCommon.IsMarkedRange(SelectionRange, ShadingColor))
                {
                    Word.Range SearchRange = SelectionRange.Duplicate;
                    while (FindNextMarkInCurrentStoryRange(ref SearchRange))
                    {
                        if (SearchRange.Start == SelectionRange.End && (bool)SearchRange.get_Information(Word.WdInformation.wdWithInTable) == (bool)SelectionRange.get_Information(Word.WdInformation.wdWithInTable))
                        {
                            SelectionRange.End = SearchRange.End;
                            if (Collapse)
                                SelectionRange.Start = SearchRange.End;

                            //if we didn't progress in this iteration, break
                            if (SelectionRange.Start == SearchRange.Start && SelectionRange.End == SearchRange.End)
                                break;
                        }
                        else
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Moves the focus to the appropriate header/footer.
        /// </summary>
        /// <param name="SelectionRange">The range containing the target header/footer.</param>
        /// <param name="sectionIdFinal">The section to be selected.</param>
        /// <param name="headerFooterId">The header/footer type to be selected.</param>
        internal static void MoveHeaderFooterFocus(Word.Range SelectionRange, int? sectionIdFinal, Word.WdHeaderFooterIndex? headerFooterId)
        {
            if (sectionIdFinal != null && headerFooterId != null)
            {
                Word.Window CurrentWindow = SelectionRange.Application.ActiveWindow;
                try
                {
                    Word.Section CurrentSection = SelectionRange.Document.Sections[(int)sectionIdFinal];
                    if (!RedactCommon.GetHeaderFooterVisibility(CurrentSection, (Word.WdHeaderFooterIndex)headerFooterId) || CurrentWindow.View.Type != Word.WdViewType.wdPrintView)
                    {
                        //if that header/footer isn't being shown, we need to fall back into normal view
                        CurrentWindow.View.Type = Word.WdViewType.wdNormalView;
                        CurrentSection.Range.Select();
                        CurrentWindow.View.SplitSpecial = RedactCommon.GetSplitTypeForStory(SelectionRange.StoryType);
                    }
                    else
                    {
                        //select the appropriate section
                        SelectionRange.Document.Sections[(int)sectionIdFinal].Range.Select();

                        //jump to the header/footer
                        CurrentWindow.View.SeekView = RedactCommon.GetSeekViewForStory(SelectionRange.StoryType);
                        
                        Thread tTabSwitch = new Thread(StayOnReviewTab);
                        tTabSwitch.Start(Word2007RedactionTool.Properties.Resources.ReviewTab);
                    }
                }
                catch (COMException)
                {
                    System.Diagnostics.Debug.Fail("incorrect header/footer view movement detected");
                }
            }
        }

        /// <summary>
        /// Move focus back to the Review tab when navigating into headers/footers.
        /// </summary>
        /// <param name="data">The localized name of the Review tab (e.g "Review" in English).</param>
        private static void StayOnReviewTab(object data)
        {
            try
            {
                TabFocusManager t = new TabFocusManager();
                t.Execute((string)data);
            }
            catch (COMException cex)
            {
                System.Diagnostics.Debug.Fail("COM Exception in tab switch: " + cex.ErrorCode);
            }
            catch (System.ComponentModel.Win32Exception wex)
            {
                System.Diagnostics.Debug.Fail("Win32 Exception in tab switch: " + wex.ErrorCode);
            }
        }
    }
}
