// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Word2007RedactionTool
{
    class RedactCommon
    {
        private RedactCommon()
        { }

        /// <summary>
        /// Gets a list of all marked ranges in the specified range. The current selection is updated to the end of the range to scan.
        /// </summary>
        /// <param name="rangeToScan">The Range to scan.</param>
        /// <param name="ShadingColor">The color for redaction marks in the document.</param>
        /// <param name="mergeAdjacent">True to merge adjacent ranges with identical formatting, False otherwise.</param>
        /// <returns>A List of RangeDataEx objects containing each marked subrange.</returns>
        internal static List<RangeDataEx> GetAllMarkedRanges(Word.Range rangeToScan, Word.WdColor ShadingColor, bool mergeAdjacent)
        {
            object Missing = Type.Missing;
            object CollapseStart = Word.WdCollapseDirection.wdCollapseStart;
            object CharacterFormatting = Word.WdUnits.wdCharacterFormatting;

            int LastPosition;
            int OriginalPosition;
            List<RangeDataEx> Ranges = new List<RangeDataEx>();
            List<RangeDataEx> RangesToDelete = new List<RangeDataEx>();
            List<RangeDataEx> ConcatenatedRanges = new List<RangeDataEx>();

            //we don't redact comments - if that's where we are, return
            if (rangeToScan.StoryType == Word.WdStoryType.wdCommentsStory)
                return ConcatenatedRanges;

            //move the selection to the beginning of the requested range
            rangeToScan.Select();

            Word.Selection CurrentSelection = rangeToScan.Application.Selection;
            CurrentSelection.Collapse(ref CollapseStart);
            OriginalPosition = CurrentSelection.Start;

            //scan for distinct ranges of formatting            
            do
            {
                //update LastPosition
                LastPosition = CurrentSelection.Start;

                //move to the next position
                CurrentSelection.Move(ref CharacterFormatting, ref Missing);

                //BUG 3913: if we detect that .Move has moved us out of the scan range (which appears to happen because of a Word bug)
                // break out and don't save the current range
                if (CurrentSelection.Start < OriginalPosition)
                {
                    CurrentSelection.End = rangeToScan.End;
                    break;
                }

                //store that range
                if (CurrentSelection.Start != LastPosition && rangeToScan.End != LastPosition)
                {
                    if (mergeAdjacent)
                    {
                        Ranges.Add(new RangeDataEx(LastPosition, CurrentSelection.Start < rangeToScan.End ? CurrentSelection.Start : rangeToScan.End, new RangeDataEx())); //since we're going to merge, don't fetch the extra properties
                    }
                    else
                    {
                        Word.Font CurrentFont = CurrentSelection.Font;
                        Ranges.Add(new RangeDataEx(LastPosition, CurrentSelection.Start < rangeToScan.End ? CurrentSelection.Start : rangeToScan.End, CurrentFont.Name, CurrentFont.Size, CurrentFont.Bold, CurrentFont.Italic, CurrentSelection.OMaths.Count > 0));
                    }
                }
            }
            while (CurrentSelection.End <= rangeToScan.End && CurrentSelection.End > LastPosition);

            if (CurrentSelection.End != rangeToScan.End)
            {
                if (mergeAdjacent)
                {
                    Ranges.Add(new RangeDataEx(CurrentSelection.End, rangeToScan.End, new RangeDataEx())); //since we're going to merge, don't fetch the extra properties
                }
                else
                {
                    Word.Font CurrentFont = CurrentSelection.Font;
                    Ranges.Add(new RangeDataEx(CurrentSelection.End, rangeToScan.End, CurrentFont.Name, CurrentFont.Size, CurrentFont.Bold, CurrentFont.Italic, CurrentSelection.OMaths.Count > 0));
                }
            }

            //go through those ranges and check if they are marked
            foreach (RangeDataEx UniqueRange in Ranges)
            {
                rangeToScan.Start = UniqueRange.Start;
                rangeToScan.End = UniqueRange.End;

                //remove the range from the list
                if (!IsMarkedRange(rangeToScan, ShadingColor))
                    RangesToDelete.Add(UniqueRange);
            }

            //clean out the list
            foreach (RangeDataEx RangeToDelete in RangesToDelete)
                Ranges.Remove(RangeToDelete);
            RangesToDelete.Clear();

            //concatenate ranges that are next to each other
            int? Start = null;
            int? End = null;
            for (int i = 0; i < Ranges.Count; i++)
            {
                //set start and end points
                if (Start == null)
                    Start = Ranges[i].Start;
                if (End == null)
                    End = Ranges[i].End;

                if ((i + 1) < Ranges.Count && (mergeAdjacent || (Ranges[i].InMath && Ranges[i + 1].InMath) || Ranges[i].IdenticalTo(Ranges[i + 1])) && End == Ranges[i + 1].Start)
                {
                    End = null;
                }
                else
                {
                    ConcatenatedRanges.Add(new RangeDataEx((int)Start, (int)End, Ranges[i]));
                    Start = End = null;
                }
            }

            //return the marked ranges
            return ConcatenatedRanges;
        }

        /// <summary>
        /// Get the Word Range for a shape, if one exists.
        /// </summary>
        /// <param name="Shape">A Shape to check for text content.</param>
        /// <returns>A Range representing the shape's content if one is present; null otherwise.</returns>
        internal static Word.Range RangeFromShape(Microsoft.Office.Interop.Word.Shape Shape)
        {
            Word.Range Range = null;

            try
            {
                if (Shape.TextFrame != null && Shape.TextFrame.HasText != 0)
                    Range = Shape.TextFrame.TextRange;
            }
            catch (UnauthorizedAccessException)
            {
                //BUG 3807 and 3849: pictures in the header will throw on the .TextFrame call, catch that
            }

            return Range;
        }

        /// <summary>
        /// Creates a copy of the current document.
        /// </summary>
        /// <param name="InputDocument">The document to clone.</param>
        /// <returns>A Document representing the cloned document.</returns>
        internal static Word.Document CloneDocument(Word.Document InputDocument)
        {
            object Missing = Type.Missing;
            object NormalTemplate = InputDocument.Application.NormalTemplate;
            object TempFile = Path.GetTempFileName();

            //copy this document into a temp file and open it
            using (StreamWriter sw = new StreamWriter((string)TempFile))
            {
                sw.Write(InputDocument.WordOpenXML);
            }
            Word.Document FileToRedact = InputDocument.Application.Documents.Add(ref TempFile, ref Missing, ref Missing, ref Missing);

            //clear the dummy template reference and move focus
            FileToRedact.set_AttachedTemplate(ref NormalTemplate);
            FileToRedact.Activate();

            //if track changes was on, turn it off
            FileToRedact.TrackRevisions = false;

            //bug 5862: lock all fields in the document body, since we shouldn't have them in the redacted copy
            FileToRedact.Fields.Unlink();

            return FileToRedact;
        }

        /// <summary>
        /// Specifies whether a range has been marked for redaction.
        /// </summary>
        /// <param name="Range">The Range to check.</param>
        /// <param name="ShadingColor">The color for redaction marks in the document.</param>
        /// <returns>True if the entire range is marked for redaction, False otherwise.</returns>
        internal static bool IsMarkedRange(Word.Range Range, Word.WdColor ShadingColor)
        {
            return Range.Font.Shading.BackgroundPatternColor == ShadingColor;
        }

        /// <summary>
        /// Create a 1px image for redacting images.
        /// </summary>
        /// <returns>A String containing the path to the 1px image.</returns>
        internal static string CreateRedactedImage()
        {
            string ImagePath = Path.GetTempFileName();
            System.Drawing.Bitmap Image = new Bitmap(1, 1);
            Image.SetPixel(0, 0, Color.Black);
            Image.Save(ImagePath);
            return ImagePath;
        }

        /// <summary>
        /// Split a range into multiple subranges, each containing exactly zero or one Shape/InlineShape.
        /// </summary>
        /// <param name="RangeToScan">A range to split.</param>
        /// <returns>A List of RangeData objects representing the subranges.</returns>
        internal static List<RangeData> SplitRange(Word.Range RangeToScan)
        {
            int i = 0;
            int LastPosition = RangeToScan.Start;
            List<RangeData> Ranges = new List<RangeData>();
            List<int> ShapeLocations = new List<int>();
            Dictionary<int, int> ShapeWidths = new Dictionary<int, int>();

            //scan for inline shapes
            foreach (Word.InlineShape InlineShape in RangeToScan.InlineShapes)
            {
                ShapeLocations.Add(InlineShape.Range.Start);
                ShapeWidths.Add(InlineShape.Range.Start, InlineShape.Range.End - InlineShape.Range.Start);
            }

            //scan for shapes
            foreach (Word.Shape Shape in RangeToScan.ShapeRange)
            {
                Word.Range AnchorRange = Shape.Anchor;
                ShapeLocations.Add(AnchorRange.Start);
                if (Shape.Type != Microsoft.Office.Core.MsoShapeType.msoCanvas)
                    ShapeWidths.Add(AnchorRange.Start, AnchorRange.End - AnchorRange.Start);
                else
                    ShapeWidths.Add(AnchorRange.Start, 28); //canvases are always 28 characters wide
            }

            //sort the anchor locations
            ShapeLocations.Sort();

            do
            {
                if (ShapeLocations[i] != LastPosition)
                    Ranges.Add(new RangeData(LastPosition, ShapeLocations[i]));
                Ranges.Add(new RangeData(ShapeLocations[i], ShapeLocations[i] + ShapeWidths[ShapeLocations[i]]));

                LastPosition = ShapeLocations[i] + ShapeWidths[ShapeLocations[i]];
                i++;
            }
            while (i < ShapeLocations.Count);

            if (LastPosition != RangeToScan.End)
                Ranges.Add(new RangeData(LastPosition, RangeToScan.End));

            return Ranges;
        }

        /// <summary>
        /// Find the previous story range (similar to Word's .NextStoryRange)
        /// </summary>
        /// <param name="SelectionRange">The starting range.</param>
        /// <returns>A Range containing the previous story range.</returns>
        internal static Word.Range PreviousStoryRange(Word.Range SelectionRange)
        {
            Word.Range PreviousRange = null;
            Word.Range CurrentRange = null;

            switch (SelectionRange.StoryType)
            {
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                case Word.WdStoryType.wdPrimaryHeaderStory:
                case Word.WdStoryType.wdEvenPagesFooterStory:
                case Word.WdStoryType.wdPrimaryFooterStory:
                case Word.WdStoryType.wdFirstPageHeaderStory:
                case Word.WdStoryType.wdFirstPageFooterStory:
                case Word.WdStoryType.wdTextFrameStory:
                    CurrentRange = SelectionRange.Document.StoryRanges[SelectionRange.StoryType];
                    while (!SelectionRange.InStory(CurrentRange) && CurrentRange.NextStoryRange != null)
                    {
                        PreviousRange = CurrentRange;
                        CurrentRange = CurrentRange.NextStoryRange;
                    }
                    return PreviousRange;
                default:
                    return null;
            }
        }

        /// <summary>
        /// Builds a string of filler characters of the same size as the input text.
        /// </summary>
        /// <param name="range">The Range for which to build filler text.</param>
        /// <param name="rangeData">The RangeDataEx containing properties about the range.</param>
        /// <returns>The filler text.</returns>
        internal static string BuildFillerText(Word.Range range, RangeDataEx rangeData)
        {
            string Result = string.Empty;
            Label Label = new Label();

            //get the correct string
            if (rangeData.Bold && rangeData.Italic)
                Label.Font = new System.Drawing.Font(rangeData.Font, rangeData.FontSize, FontStyle.Italic | FontStyle.Bold);
            else if (rangeData.Bold)
                Label.Font = new System.Drawing.Font(rangeData.Font, rangeData.FontSize, FontStyle.Bold);
            else if (rangeData.Italic)
                Label.Font = new System.Drawing.Font(rangeData.Font, rangeData.FontSize, FontStyle.Italic);
            else
                Label.Font = new System.Drawing.Font(rangeData.Font, rangeData.FontSize, FontStyle.Regular);

            Label.AutoSize = true;

            char[] delim = { ' ', '\v', '\f' }; // \v is a soft break, \f is a page break
            string text = range.Text;
            if (text == null)
                return null;

            string[] lines = text.Split(delim);

            int c = 0;
            Random rnd = new Random();
            foreach (string line in lines)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    Label.Text = line;                    
                    float obfuscationFactor = (float)rnd.Next(90, 110) / 100; // make each word +/- 10% to make it harder to determine the original value
                    float originalWidth = Label.PreferredWidth * obfuscationFactor;
                    for (Label.Text = string.Empty; Label.PreferredWidth < originalWidth; Label.Text += "'") 
                        ;
                    Result += Label.Text;
                    c += line.Length;
                }

                if (c < text.Length)
                {
                    Result += text[c];
                    c++;
                }
            }
            return Result;
        }

        /// <summary>
        /// Gets the SeekView constant needed to access a given document story.
        /// </summary>
        /// <param name="Story">The document story to be accessed.</param>
        /// <returns>The corresponding SeekView constant.</returns>
        internal static Word.WdSeekView GetSeekViewForStory(Word.WdStoryType Story)
        {
            switch (Story)
            {
                case Word.WdStoryType.wdEvenPagesFooterStory:
                    return Word.WdSeekView.wdSeekEvenPagesFooter;
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                    return Word.WdSeekView.wdSeekEvenPagesHeader;
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return Word.WdSeekView.wdSeekFirstPageFooter;
                case Word.WdStoryType.wdFirstPageHeaderStory:
                    return Word.WdSeekView.wdSeekFirstPageHeader;
                case Word.WdStoryType.wdPrimaryFooterStory:
                    return Word.WdSeekView.wdSeekPrimaryFooter;
                case Word.WdStoryType.wdPrimaryHeaderStory:
                    return Word.WdSeekView.wdSeekPrimaryHeader;
                default:
                    return Word.WdSeekView.wdSeekMainDocument;
            }
        }

        /// <summary>
        /// Gets whether or not a specific header/footer type is visible in a given section.
        /// </summary>
        /// <param name="Section">The section to check.</param>
        /// <param name="HeaderFooterIndex">The header/footer to check for.</param>
        /// <returns>True if visible, otherwise False.</returns>
        internal static bool GetHeaderFooterVisibility(Word.Section Section, Word.WdHeaderFooterIndex HeaderFooterIndex)
        {
            Word.Range StartOfSection = Section.Range.Duplicate;
            Word.Range EndOfSection = Section.Range.Duplicate;
            
            //set ranges
            StartOfSection.End = StartOfSection.Start;

            //get start/end page
            int StartingPageNumber = (int)StartOfSection.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
            int EndingPageNumber = (int)EndOfSection.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);

            //first page header in use?
            bool FirstPageVisible = (Section.PageSetup.DifferentFirstPageHeaderFooter == -1);
            if (FirstPageVisible && HeaderFooterIndex == Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                return true;

            //even page header in use?
            bool EvenPagesInUse = (Section.PageSetup.OddAndEvenPagesHeaderFooter == -1);
            if (!EvenPagesInUse && HeaderFooterIndex == Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                return false;

            if (EndingPageNumber - StartingPageNumber >= 2)
            {
                //if there are more than three pages, then the type is definitely showing
                return true;
            }
            else if (EndingPageNumber - StartingPageNumber == 1)
            {
                //if there are two pages, need to check
                switch (HeaderFooterIndex)
                {
                    case Word.WdHeaderFooterIndex.wdHeaderFooterPrimary:
                        if (!EvenPagesInUse || !FirstPageVisible)
                            return true;
                        else
                            return ((StartingPageNumber + 1) % 2 != 0); // is next page odd
                    case Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages:
                        if (!FirstPageVisible)
                            return true;
                        else
                            return ((StartingPageNumber + 1) % 2 == 0); // is next page even
                }
            }            
            
            // == 1
            switch (HeaderFooterIndex)
            {
                case Word.WdHeaderFooterIndex.wdHeaderFooterPrimary:
                    if (FirstPageVisible) // first page is showing
                        return false;
                    else if (EvenPagesInUse) // first page not showing and even showing
                        return (StartingPageNumber % 2 != 0);
                    else //first page not showing, even not showing
                        return true; // is page odd
                case Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages:
                    if (FirstPageVisible) // first page is showing
                        return false;
                    else if (EvenPagesInUse) // first page not showing and even showing
                        return (StartingPageNumber % 2 == 0);
                    else //first page not showing, even not showing
                        return false; // is page odd
            }

            Debug.Fail("should never fall through on visibility");
            return false;
        }

        /// <summary>
        /// Finds the index of the section with the specified header/footer.
        /// </summary>
        /// <param name="Document">The document to search.</param>
        /// <param name="SelectionRange">The range containing the specified header/footer.</param>
        /// <returns>The ordinal position of the section containing the header/footer.</returns>
        internal static int FindSectionWithHeaderFooter(Word.Document Document, Word.Range SelectionRange)
        {
            for (int i = 1; i <= Document.Sections.Count; i++)
            {
                if (IsHeader(SelectionRange.StoryType))
                {
                    if (SelectionRange.InStory(Document.Sections[i].Headers[GetHeaderFooterType(SelectionRange.StoryType)].Range))
                        return i;
                }
                else
                {
                    if (SelectionRange.InStory(Document.Sections[i].Footers[GetHeaderFooterType(SelectionRange.StoryType)].Range))
                        return i;
                }
            }

            throw new ArgumentException("could not locate story for range");
        }

        /// <summary>
        /// Gets a HeaderFooterIndex from a WdStoryType.
        /// </summary>
        /// <param name="StoryType">The StoryType to convert.</param>
        /// <returns>The type of HeaderFooterIndex corresponding to that StoryType.</returns>
        /// <exception cref="ArgumentException">If the story type is not a valid header/footer story type, throws.</exception>
        internal static Word.WdHeaderFooterIndex GetHeaderFooterType(Word.WdStoryType StoryType)
        {
            switch (StoryType)
            {
                case Word.WdStoryType.wdPrimaryHeaderStory:
                case Word.WdStoryType.wdPrimaryFooterStory:
                    return Word.WdHeaderFooterIndex.wdHeaderFooterPrimary;
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                case Word.WdStoryType.wdEvenPagesFooterStory:
                    return Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages;
                case Word.WdStoryType.wdFirstPageHeaderStory:
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage;
                default:
                    throw new ArgumentException("this is a non-header/footer story type");
            }
        }

        /// <summary>
        /// Determines whether the input story is a header or a footer.
        /// </summary>
        /// <param name="StoryType">The story type to check.</param>
        /// <returns>True if it is a header, False if it is a footer.</returns>
        /// <exception cref="ArgumentException">If the story type is not a valid header/footer story type, throws.</exception>
        private static bool IsHeader(Word.WdStoryType StoryType)
        {
            switch (StoryType)
            {
                case Word.WdStoryType.wdPrimaryHeaderStory:
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                case Word.WdStoryType.wdFirstPageHeaderStory:
                    return true;                
                case Word.WdStoryType.wdPrimaryFooterStory:
                case Word.WdStoryType.wdEvenPagesFooterStory:
                case Word.WdStoryType.wdFirstPageFooterStory: 
                    return false;
                default:
                    throw new ArgumentException("this is a non header/footer story type");
            }
        }

        /// <summary>
        /// Determines whether the specified story is a header or footer.
        /// </summary>
        /// <param name="StoryType">The story type to check.</param>
        /// <returns>True if it is a header or footer, otherwise False.</returns>
        internal static bool IsHeaderFooter(Word.WdStoryType StoryType)
        {
            switch (StoryType)
            {
                case Word.WdStoryType.wdPrimaryHeaderStory:
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                case Word.WdStoryType.wdFirstPageHeaderStory:
                case Word.WdStoryType.wdPrimaryFooterStory:
                case Word.WdStoryType.wdEvenPagesFooterStory:
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Gets a WdSpecialPane from a WdStoryType.
        /// </summary>
        /// <param name="StoryType">The StoryType to convert.</param>
        /// <returns>The WdSpecialPane corresponding to that StoryType.</returns>
        internal static Word.WdSpecialPane GetSplitTypeForStory(Word.WdStoryType StoryType)
        {
            switch (StoryType)
            {
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                    return Word.WdSpecialPane.wdPaneEvenPagesHeader;
                case Word.WdStoryType.wdEvenPagesFooterStory:
                    return Word.WdSpecialPane.wdPaneEvenPagesFooter;
                case Word.WdStoryType.wdFirstPageHeaderStory:
                    return Word.WdSpecialPane.wdPaneFirstPageHeader;
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return Word.WdSpecialPane.wdPaneFirstPageFooter;
                case Word.WdStoryType.wdPrimaryHeaderStory:
                    return Word.WdSpecialPane.wdPanePrimaryHeader;
                case Word.WdStoryType.wdPrimaryFooterStory:
                    return Word.WdSpecialPane.wdPanePrimaryFooter;
                default:
                    throw new ArgumentException("unexpected story type");
            }
        }

        /// <summary>
        /// Gets the document surface window.
        /// </summary>
        /// <returns>A NativeWindow representing the document surface window.</returns>
        internal static NativeWindow GetDocumentSurfaceWindow()
        {
            NativeWindow WordWindow = new NativeWindow();

            //traverse down to get the window for the document surface
            IntPtr WindowPtr = NativeMethods.FindWindowByClass("OpusApp", IntPtr.Zero);
            WindowPtr = NativeMethods.FindWindowEx(WindowPtr, IntPtr.Zero, "_WwF", IntPtr.Zero);
            WindowPtr = NativeMethods.FindWindowEx(WindowPtr, IntPtr.Zero, "_WwB", IntPtr.Zero);
            WindowPtr = NativeMethods.FindWindowEx(WindowPtr, IntPtr.Zero, "_WwG", IntPtr.Zero);
            System.Diagnostics.Debug.WriteLine(WindowPtr.ToString());

            WordWindow.AssignHandle(WindowPtr);

            return WordWindow;
        }
    }

    /// <summary>
    /// A class defining a localizable message box.
    /// </summary>
    public static class GenericMessageBox
    {
        /// <summary>
        /// Displays a dialog box using the specified parameters.
        /// </summary>
        /// <param name="text">The text to display.</param>
        /// <param name="caption">The caption to display.</param>
        /// <param name="buttons">The buttons to display.</param>
        /// <param name="icon">The icon to display.</param>
        /// <param name="defaultButton">The default button for the dialog box.</param>
        /// <param name="options">A MessageBoxOptions object specifying an other options.</param>
        /// <returns>A DialogResult corresponding to the user's action on the dialog box.</returns>
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options)
        {
            if (IsRightToLeft())
            {
                options |= MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign;
            }

            return MessageBox.Show(text, caption, buttons, icon, defaultButton, options);
        }

        /// <summary>
        /// Check if Word is running right-to-left.
        /// </summary>
        /// <returns>True if RTL, False if LTR.</returns>
        private static bool IsRightToLeft()
        {
            return CultureInfo.GetCultureInfo((int)Globals.ThisAddIn.Application.Language).TextInfo.IsRightToLeft;
        }
    }

    /// <summary>
    /// A struct defining a range from its start and end positions.
    /// </summary>
    internal struct RangeData
    {
        private int m_Start;
        private int m_End;

        public RangeData(int rangeStart, int rangeEnd)
        {
            m_Start = rangeStart;
            m_End = rangeEnd;
        }

        public int Start
        {
            get { return m_Start; }
        }

        public int End
        {
            get { return m_End; }
        }
    }

    /// <summary>
    /// A struct defining a range from its start/end positions and formatting properties.
    /// </summary>
    internal struct RangeDataEx
    {
        private int m_Start;
        private int m_End;
        private string m_Font;
        private float m_FontSize;
        private bool m_Bold;
        private bool m_Italics;
        private bool m_InMath;

        public RangeDataEx(int rangeStart, int rangeEnd, string Font, float FontSize, int Bold, int Italics, bool InMath)
        {
            m_Start = rangeStart;
            m_End = rangeEnd;
            m_Font = Font;
            m_FontSize = FontSize;
            m_Bold = (Bold == -1);
            m_Italics = (Italics == -1);
            m_InMath = InMath;
        }

        /// <summary>
        /// Creates a new RangeDataEx object.
        /// </summary>
        /// <param name="rangeStart">The starting point of the specified range.</param>
        /// <param name="rangeEnd">The ending point of the specified range.</param>
        /// <param name="baseRange">An existing RangeDataEx from which to clone formatting information.</param>
        public RangeDataEx(int rangeStart, int rangeEnd, RangeDataEx baseRange)
        {
            m_Start = rangeStart;
            m_End = rangeEnd;
            m_Font = baseRange.m_Font;
            m_FontSize = baseRange.m_FontSize;
            m_Bold = baseRange.m_Bold;
            m_Italics = baseRange.m_Italics;
            m_InMath = baseRange.m_InMath;
        }

        public int Start
        {
            get { return m_Start; }
        }

        public int End
        {
            get { return m_End; }
        }

        public string Font
        {
            get { return m_Font; }
        }

        public float FontSize
        {
            get { return m_FontSize; }
        }

        public bool Bold
        {
            get { return m_Bold; }
        }

        public bool Italic
        {
            get { return m_Italics; }
        }

        public bool InMath
        {
            get { return m_InMath; }
        }

        /// <summary>
        /// Determines if two ranges have the same formatting applied to them.
        /// </summary>
        /// <param name="ComparisonRange">A RangeDataEx to compare to.</param>
        /// <returns>True if identical, False otherwise.</returns>
        public bool IdenticalTo(RangeDataEx ComparisonRange)
        {
            if (m_Bold != ComparisonRange.m_Bold)
                return false;
            else if (m_Italics != ComparisonRange.m_Italics)
                return false;
            else if (m_Font != ComparisonRange.m_Font)
                return false;
            else if (m_FontSize != ComparisonRange.m_FontSize)
                return false;
            else
                return true;
        }
    }
}
