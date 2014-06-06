using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Reflection;
//using NLog.Targets.Wrappers;
//using NLog;
//using NLog.Targets;
using System.IO;
using System.Diagnostics;
using System.Net;
namespace MSWordReportingEditor
{
    public partial class WordControl : UserControl
    {
        //private static Logger logger = LogManager.GetLogger("WordControl.cs");

        #region "Password Protect"
        object _noReset = false;
        object _password = "ANGAD";
        object _useIRM = false;
        object _enforceStyleLock = false;
        Bookmark _userReportBookmark;
        Bookmark _userImpressionBookmark;
        Bookmark _userAddendumTitle;
        Bookmark _userAddendumText;


        #endregion
        #region "API usage declarations"

        [DllImport("user32.dll")]
        public static extern int FindWindow(string strclassName, string strWindowName);

        [DllImport("user32.dll")]
        static extern int SetParent(int hWndChild, int hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        static extern bool SetWindowPos(
            int hWnd,               // handle to window
            int hWndInsertAfter,    // placement-order handle
            int X,                  // horizontal position
            int Y,                  // vertical position
            int cx,                 // width
            int cy,                 // height
            uint uFlags             // window-positioning options
            );

        [DllImport("user32.dll", EntryPoint = "MoveWindow")]
        static extern bool MoveWindow(
            int hWnd,
            int X,
            int Y,
            int nWidth,
            int nHeight,
            bool bRepaint
            );

        [DllImport("user32.dll", EntryPoint = "DrawMenuBar")]
        static extern Int32 DrawMenuBar(
            Int32 hWnd
            );

        [DllImport("user32.dll", EntryPoint = "GetMenuItemCount")]
        static extern Int32 GetMenuItemCount(
            Int32 hMenu
            );

        [DllImport("user32.dll", EntryPoint = "GetSystemMenu")]
        static extern Int32 GetSystemMenu(
            Int32 hWnd,
            bool bRevert
            );

        [DllImport("user32.dll", EntryPoint = "RemoveMenu")]
        static extern Int32 RemoveMenu(
            Int32 hMenu,
            Int32 nPosition,
            Int32 wFlags
            );


        private const int MF_BYPOSITION = 0x400;
        private const int MF_REMOVE = 0x1000;


        const int SWP_DRAWFRAME = 0x20;
        const int SWP_NOMOVE = 0x2;
        const int SWP_NOSIZE = 0x1;
        const int SWP_NOZORDER = 0x4;

        #endregion

        //Represent word document
        public Document _wordDocument;
        //Represent word application
        public static ApplicationClass _wordApp = null;
        //Represent application IntPtr handle.
        public static int _hWnd = 0;
        //Represent a file that will be loaded.
        public static string _inpFilePath = null;
        //I don't know why is this variable but it's definetly used in code.
        private static bool deactivateevents = false;

        private string _wordDocumentCaption = null;
        //
        private static string _blankTemplatePath;

        private string[] macros = { "$(P_NAME)", "$(P_ID)", "$(ACC_NO)", "$(MODALITY)", "$(SCAN_DATE)", "$(SCAN_TIME)", "$(P_SEX)",
                                    "$(P_AGE)", "$(REF_PHYS)", "$(PROCEDURE_DESC)","$(REF_DEPT)","$(DATE/TIME)","$(CLINICAL_DETAILS)",
                                  "$(ORDER_NUMBER)","$(STUDY_PART)","$(LOCATION)","$(EXAMINATION_DESC)",
                                  "$(P_DOB)","$(ITEM_DESC)","$(ORGAN)","$(REPORTED_DR)","$(ROUTED_DR)","$(APPROVED_DR)",
                                  "$(TYPED_DR)","$(APPROVED_DATE/TIME)"};

        public string reportedDateAndTime = "";        
        public DateTime ServerDepReportDateTime;

        private int blankTemplateWordCount = 0;
        
        //static readonly string textMarkerInvisible = "qqqzzz";
        /// <summary>
        /// User control that contains Embedded MS WORD 
        /// </summary>
        public WordControl()
        {
            InitializeComponent();
            SizeChanged += new EventHandler(WordControl_SizeChanged);


            this.InitializeWordApp();
            //this.timerAutoSave.Enabled = App.AutoSaveEnabled;
#if DEBUG
            //this.timerAutoSave.Interval = 10000; // We can't so much for the development.
#endif 
        }

        void WordControl_SizeChanged(object sender, EventArgs e)
        {

        }

        //public ReportAdditionalInformation AdditionalInformation
        //{
        //    get { return JSONStore.GetDataObject; }
        //}

        public string BlankTemplateFilePath
        {
            get { return _blankTemplatePath; }
            set { _blankTemplatePath = value; }
        }
        /// <summary>
        /// This code is to reopen Word.
        /// </summary>
        public void CloseControl()
        {
            if (_wordApp == null)
            {
                //logger.Info("No WINWORD application to close.");
                return;
            }
            try
            {
                deactivateevents = true;
                object dummy = null;
                object dummy2 = (object)false;
                CloseDocument();
                // Change the line below.
                _wordApp.Quit(ref dummy2, ref dummy, ref dummy);
                deactivateevents = false;
            //    HttpWebResponse _webResponse = null;
            //    string url = App.ServerProtocol + App.ServerIP + ":" + Convert.ToString(App.ServerPort) + "/MSWORD/removeReportLocks?sessionId=" + App.SessionId + "&amp;reportKey=" + App.ReportKey;
            //    HttpWebRequest _webRequest = (HttpWebRequest)WebRequest.Create(url);
            //    Cookie sessionCookie = new Cookie();
            //    sessionCookie.HttpOnly = true;
            //    sessionCookie.Domain = App.ServerIP;
            //    sessionCookie.Name = "session_id";
            //    sessionCookie.Value = App.SessionId;
            //    _webRequest.CookieContainer = new CookieContainer();
            //    _webRequest.CookieContainer.Add(sessionCookie);
            //    _webRequest.Method = "GET";
            //    _webResponse = (HttpWebResponse)_webRequest.GetResponse();
            }
            catch (Exception ex)
            {
                //logger.ErrorException("<Exception>" + ex.ToString(), ex);
                //MessageBox.Show(e.ToString());
            }
        }

        private void OnClose(Document doc, ref bool cancel)
        {
            if (!deactivateevents)
            {
                cancel = true;
            }
        }

        private void OnOpenDoc(Document doc)
        {
            OnNewDoc(doc);
        }

        private void OnNewDoc(Document doc)
        {
            if (!deactivateevents)
            {
                deactivateevents = true;
                object dummy = null;
                object dummy2 = (object)false;

                // Change the line below.
                _wordApp.Quit(ref dummy2, ref dummy, ref dummy);
                deactivateevents = false;
            }
        }

        private void OnQuit()
        {
            //_wordApp=null;
        }

        private void AddHeader(bool inEveryPage)
        {
            if (_wordDocument == null)
            {
                //logger.Fatal("No active document found. Unable to add Patient demographics header");
                return;
            }
            _wordApp.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = -1;
            object start = 0;
            object end = 0;
            object oNull = Missing.Value;
            Range _otherPageHeader = null;

            Range _firstPagedHeader = null;
            if (inEveryPage)
            {
                // reference: http://msdn.microsoft.com/en-us/library/ms178795(v=vs.80).aspx
                foreach (Section section in _wordDocument.Sections)
                {
                    _firstPagedHeader = section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    _otherPageHeader = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            
                }
            }
            else
            {
                _otherPageHeader = _wordDocument.Range(ref start, ref end);
            }
            _firstPagedHeader.Font.Name = _otherPageHeader.Font.Name = "Times New Roman";
            _firstPagedHeader.Font.Size = _otherPageHeader.Font.Size = 13F;
            MoveCursorToEndOfDocument();
            // how to add multiple tables in word.interop.range
            // ref: http://stackoverflow.com/questions/6996242/how-adding-multi-tables-on-word-with-net

            // some how top margin did not took effect in header section. 
            // therefore adding a 1x1 table with height as top margin equivalent. 

            //add 1x1 table
            Table tbl1 = _wordDocument.Tables.Add(_firstPagedHeader, 1, 1, ref oNull, ref oNull);
            Table tbl2 = _wordDocument.Tables.Add(_otherPageHeader, 1, 1, ref oNull, ref oNull);
            // set row height
            tbl1.Cell(1, 1).Row.Height = _wordApp.InchesToPoints((float)Convert.ToDouble(JSONStore.GetDataObject.topMargin));
            tbl2.Cell(1, 1).Row.Height = _wordApp.InchesToPoints((float)Convert.ToDouble(JSONStore.GetDataObject.topMargin));

            object collapseDirection = WdCollapseDirection.wdCollapseEnd;
            // move cursor to the end of table.
            _firstPagedHeader.Collapse(ref collapseDirection);
            _otherPageHeader.Collapse(ref collapseDirection);
            // Now add something behind the table to prevent word from joining tables into one
            _firstPagedHeader.InsertParagraphAfter();
            _otherPageHeader.InsertParagraphAfter();
            // don't know why this is needed. but reference link told so. 
            // need to move to the end again
            _firstPagedHeader.Collapse(ref collapseDirection);
            _otherPageHeader.Collapse(ref collapseDirection);


            Table firstPageTable = _wordDocument.Tables.Add(_firstPagedHeader, 5, 6, ref oNull, ref oNull);
            Table otherPageTable = _wordDocument.Tables.Add(_otherPageHeader, 1, 9, ref oNull, ref oNull);

            firstPageTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            firstPageTable.Borders.OutsideColor = WdColor.wdColorBlack;
            firstPageTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            firstPageTable.Borders.InsideColor = WdColor.wdColorBlack;

            otherPageTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            otherPageTable.Borders.OutsideColor = WdColor.wdColorBlack;
            otherPageTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            otherPageTable.Borders.InsideColor = WdColor.wdColorBlack;


            firstPageTable.Columns[1].SetWidth(_wordApp.InchesToPoints(1.56f), WdRulerStyle.wdAdjustNone);
            firstPageTable.Columns[2].SetWidth(_wordApp.InchesToPoints(0.19f), WdRulerStyle.wdAdjustNone);
            firstPageTable.Columns[3].SetWidth(_wordApp.InchesToPoints(2.38f), WdRulerStyle.wdAdjustNone);


            firstPageTable.Columns[4].SetWidth(_wordApp.InchesToPoints(1.61f), WdRulerStyle.wdAdjustNone);
            firstPageTable.Columns[5].SetWidth(_wordApp.InchesToPoints(0.19f), WdRulerStyle.wdAdjustNone);
            firstPageTable.Columns[6].SetWidth(_wordApp.InchesToPoints(1.76f), WdRulerStyle.wdAdjustNone);
            
            otherPageTable.Columns[1].SetWidth(_wordApp.InchesToPoints(0.63f), WdRulerStyle.wdAdjustNone);
            otherPageTable.Columns[2].SetWidth(_wordApp.InchesToPoints(0.19f), WdRulerStyle.wdAdjustNone);
            otherPageTable.Columns[3].SetWidth(_wordApp.InchesToPoints(2.62f), WdRulerStyle.wdAdjustNone);


            otherPageTable.Columns[4].SetWidth(_wordApp.InchesToPoints(0.36f), WdRulerStyle.wdAdjustNone);
            otherPageTable.Columns[5].SetWidth(_wordApp.InchesToPoints(0.19f), WdRulerStyle.wdAdjustNone);
            otherPageTable.Columns[6].SetWidth(_wordApp.InchesToPoints(1.43f), WdRulerStyle.wdAdjustNone);

            otherPageTable.Columns[7].SetWidth(_wordApp.InchesToPoints(0.77f), WdRulerStyle.wdAdjustNone);
            otherPageTable.Columns[8].SetWidth(_wordApp.InchesToPoints(0.19f), WdRulerStyle.wdAdjustNone);
            otherPageTable.Columns[9].SetWidth(_wordApp.InchesToPoints(1.23f), WdRulerStyle.wdAdjustNone);
            

            firstPageTable.AllowAutoFit = true;
            firstPageTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

            otherPageTable.AllowAutoFit = true;
            otherPageTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

            firstPageTable.Rows[1].Cells.Merge();
            firstPageTable.Cell(1, 1).Range.Text = "Department Of Molecular Imaging & Radiology" + Environment.NewLine;
            firstPageTable.Cell(1, 1).Range.Font.Name = "Times New Roman";
            firstPageTable.Cell(1, 1).Range.Font.Size = 15;
            firstPageTable.Cell(1, 1).Range.Bold = 1;
            firstPageTable.Cell(1, 1).Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
            firstPageTable.Cell(1, 1).Range.Underline = WdUnderline.wdUnderlineSingle;
            firstPageTable.Cell(1, 1).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


            firstPageTable.Cell(2, 1).Range.Text = "Name";
            firstPageTable.Cell(2, 1).Range.Bold = 1;
            firstPageTable.Cell(2, 1).Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            firstPageTable.Cell(2, 2).Range.Text = ":";
            firstPageTable.Cell(2, 2).Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            firstPageTable.Cell(2, 3).Range.Text = JSONStore.GetDataObject.patientData.patientname;
            firstPageTable.Cell(2, 3).Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            firstPageTable.Cell(2, 4).Range.Text = "Id";
            firstPageTable.Cell(2, 4).Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            firstPageTable.Cell(2, 4).Range.Bold = 1;
            firstPageTable.Cell(2, 5).Range.Text = ":";
            firstPageTable.Cell(2, 5).Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            firstPageTable.Cell(2, 6).Range.Text = JSONStore.GetDataObject.patientData.patientid;
            firstPageTable.Cell(2, 6).Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;


            firstPageTable.Cell(3, 1).Range.Text = "Age/Sex";
            firstPageTable.Cell(3, 1).Range.Bold = 1;
            firstPageTable.Cell(3, 2).Range.Text = ":";
            firstPageTable.Cell(3, 3).Range.Text = JSONStore.GetDataObject.patientData.date_of_birth + "/" + JSONStore.GetDataObject.patientData.sex;

            firstPageTable.Cell(3, 4).Range.Text = "Accession no.";
            firstPageTable.Cell(3, 4).Range.Bold = 1;
            firstPageTable.Cell(3, 5).Range.Text = ":";
            firstPageTable.Cell(3, 6).Range.Text = JSONStore.GetDataObject.patientData.accessionNo;

            
            firstPageTable.Cell(4, 1).Range.Text = "Modality";
            firstPageTable.Cell(4, 1).Range.Bold = 1;
            firstPageTable.Cell(4, 2).Range.Text = ":";
            firstPageTable.Cell(4, 3).Range.Text = JSONStore.GetDataObject.patientData.modality;

            firstPageTable.Cell(4, 4).Range.Text = "Report Date/Time";
            firstPageTable.Cell(4, 4).Range.Bold = 1;
            firstPageTable.Cell(4, 5).Range.Text = ":";
            firstPageTable.Cell(4, 6).Range.Text = string.Empty;


            firstPageTable.Cell(5, 1).Range.Text = "Ref. Doctor";
            firstPageTable.Cell(5, 1).Range.Bold = 1;
            firstPageTable.Cell(5, 2).Range.Text = ":";
            firstPageTable.Cell(5, 3).Range.Text = JSONStore.GetDataObject.patientData.doctorname;

            firstPageTable.Cell(5, 4).Range.Text = "Scan Date/Time";
            firstPageTable.Cell(5, 4).Range.Bold = 1;
            firstPageTable.Cell(5, 5).Range.Text = ":";
            firstPageTable.Cell(5, 6).Range.Text = JSONStore.GetDataObject.patientData.studydate + "/" + JSONStore.GetDataObject.patientData.studytime;

            otherPageTable.Cell(1, 1).Range.Text = "Name";
            otherPageTable.Cell(1, 1).Range.Bold = 1;
            otherPageTable.Cell(1, 2).Range.Text = ":";
            otherPageTable.Cell(1, 3).Range.Text = JSONStore.GetDataObject.patientData.patientname;

            otherPageTable.Cell(1, 4).Range.Text = "Id";
            otherPageTable.Cell(1, 4).Range.Bold = 1;
            otherPageTable.Cell(1, 5).Range.Text = ":";
            otherPageTable.Cell(1, 6).Range.Text = JSONStore.GetDataObject.patientData.patientid;

            otherPageTable.Cell(1, 7).Range.Text = "Age/Sex";
            otherPageTable.Cell(1, 7).Range.Bold = 1;
            otherPageTable.Cell(1, 8).Range.Text = ":";
            otherPageTable.Cell(1, 9).Range.Text = JSONStore.GetDataObject.patientData.date_of_birth + "/" + JSONStore.GetDataObject.patientData.sex;
        }

        private void AddHeader1()
        {
            object start = 0;
            object end = 0;
            Range _otherPageHeader = null;

            Range _firstPagedHeader = null;
            foreach (Section section in _wordDocument.Sections)
            {
                _firstPagedHeader = section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                InitAndUpdateMergeFields(ref _firstPagedHeader);
                Range _firstPageFooter = section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                InitAndUpdateMergeFields(ref _firstPageFooter);

                _otherPageHeader = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                InitAndUpdateMergeFields(ref _otherPageHeader);
                
                Range _otherPageFooter = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                InitAndUpdateMergeFields(ref _otherPageFooter);

                _wordDocument.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;

            }
            Range r = _wordDocument.Range(ref start, ref end) ;
            InitAndUpdateMergeFields(ref r);

        }

        private void InitAndUpdateMergeFields(ref Range range)
        {
            if (_wordDocument == null)
            {
                //logger.Fatal("Word Document has not been initialized.");
                return;
            }

            PreviousReport pr = JSONStore.GetDataObject.getPreviousReport(JSONStore.GetDataObject.latestReportId);
            foreach (string macro in macros)
            {
                switch (macro)
                {
                    case "$(P_NAME)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.patientname; 
                        break;

                    case "$(P_ID)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.patientid; 
                        break;

                    case "$(ACC_NO)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.accessionNo; 
                        break;

                    case "$(MODALITY)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.modality;
                        break;
                    case "$(SCAN_DATE)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.studydate;
                        break;

                    case "$(SCAN_TIME)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.studytime;
                        break;

                    case "$(P_SEX)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.sex;
                        break;
                    case "$(P_AGE)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.age;
                        break;
                    case "$(REF_PHYS)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.doctorname;
                        break;
                    case "$(ORDER_NUMBER)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.displayOrderNumber;
                        break;
                    case "$(ITEM_DESC)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.procedureDescription;
                        break;
                    case "$(ORGAN)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.organ;
                        break;
                    case "$(P_DOB)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.date_of_birth;
                        break;
                    case "$(REF_DEPT)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.refPhyDept;
                        break;
                    case "$(CLINICAL_DETAILS)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.clinicalDetails;
                        break;
                    case "$(LOCATION)":
                        range.Find.Text = macro;
                        range.Find.Replacement.Text = JSONStore.GetDataObject.patientData.location;
                        break;
                    case "$(REPORTED_DR)":
                        if (App.ApplicationMode == ApplicationMode.RadiologistMode)
                        {
                            range.Find.Text = macro;
                            range.Find.Replacement.Text = JSONStore.GetDataObject.userData.fname + " " + JSONStore.GetDataObject.userData.mname + " " + JSONStore.GetDataObject.userData.lname;
                            range.Find.Replacement.Font.Color = WdColor.wdColorAutomatic;
                        }                        
                        break;
                    case "$(TYPED_DR)":
                        if (App.ApplicationMode == ApplicationMode.TypistMode)
                        {
                            range.Find.Text = macro;
                            range.Find.Replacement.Text = JSONStore.GetDataObject.userData.fname + " " + JSONStore.GetDataObject.userData.mname + " " + JSONStore.GetDataObject.userData.lname;
                            range.Find.Replacement.Font.Color = WdColor.wdColorAutomatic;
                        }
                        break;
                    case "$(DATE/TIME)":
                        range.Find.Text = macro;
                        if (JSONStore.GetDataObject.latestReportId < 0)
                            reportedDateAndTime = Utils.MyDateTimeFormat.GetCurrentDateTimeFormattedForReport();
                        else
                            reportedDateAndTime = Utils.MyDateTimeFormat.GetDateTimeFormattedForReport(pr.date, pr.time);
                        range.Find.Replacement.Text = reportedDateAndTime;
                        break;

                    default:
                        break;
                }
                object replaceAll = WdReplace.wdReplaceAll;
                object missing = System.Type.Missing;
            
                // Execute the Find and Replace -- notice that the
                // 11th parameter is the "replaceAll" enum object
                bool res =range.Find.Execute(ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref replaceAll,
                    ref missing, ref missing, ref missing, ref missing);
             
            }
        }

        /// <summary>
        /// Method to update merge fields with appropriate values.
        /// </summary>
        
        private void MoveCursorToEndOfDocument()
        {
            //object missing = Missing.Value;
            //object what = WdGoToItem.wdGoToLine;
            //object which = WdGoToDirection.wdGoToLast;
            //_wordApp.Selection.GoTo(ref what, ref which, ref missing, ref missing);

            //object missing = Missing.Value;
            //object what = WdGoToItem.wdGoToPercent;
            //object which = WdGoToDirection.wdGoToLast;
            //_wordApp.Selection.GoTo(ref what, ref which, ref missing, ref missing);

            object bookmarkName = "REPORT_TEXT";
            if (_wordDocument.Bookmarks.Exists(bookmarkName.ToString()))
            {
                Bookmark bookmark = _wordDocument.Bookmarks.get_Item(ref bookmarkName);
                bookmark.Select();
            }
            
            //object wdStory = WdUnits.wdStory;
            //object extended = WdMovementType.wdMove;
            //_wordApp.Selection.EndKey(ref wdStory, ref extended);
            
        }
        private void MoveCursorToStartOfDocument()
        {
            object missing = Missing.Value;
            object what = WdGoToItem.wdGoToPercent;
            object which = WdGoToDirection.wdGoToFirst;
           
            _wordApp.Selection.GoTo(ref what, ref which, ref missing, ref missing);
        }

        public void CloseDocument()
        {
            object dummy = null;
            object saveChanges = (object)false;
            if (_wordDocument == null)
            {
                //logger.Info("No document present to close");
            }
            else
            {
                try
                {
                    string filePath = _wordDocument.FullName;
                    _wordDocument.Close(ref saveChanges, ref dummy, ref dummy);                    
                    _wordDocument = null;
                    File.Delete(filePath);
                   
                }
                catch (Exception ex)
                {
                    //logger.ErrorException("CloseDocument ",ex);
                }
            }

        }

        public void FindAndReplaceAll(ref Range range,string textToFind,string textToReplaceWith)
        {
            //object EditAccess = WdEditorType.wdEditorEveryone;

            //range.Editors.Add(ref EditAccess);

            object replaceAll = WdReplace.wdReplaceAll;
                
            object missing = System.Type.Missing;

            range.Find.ClearFormatting();
            //range.Find.MatchWildcards = true;
            range.Find.Text = textToFind;
            range.Find.Replacement.Text = textToReplaceWith;
            

            // Execute the Find and Replace -- notice that the
            // 11th parameter is the "replaceAll" enum object
           bool res= range.Find.Execute(ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref replaceAll,
                ref missing, ref missing, ref missing, ref missing);
        }
        //public void FindAndReplace(object findText, object replaceWithText)
        //{
        //    object matchCase = true;
        //    object matchWholeWord = true;
        //    object matchWildCards = false;
        //    object matchSoundsLike = false;
        //    object nmatchAllWordForms = false;
        //    object forward = true;
        //    object format = false;
        //    object matchKashida = false;
        //    object matchDiacritics = false;
        //    object matchAlefHamza = false;
        //    object matchControl = false;
        //    object read_only = false;
        //    object visible = true;
        //    object replace = 2;
        //    object wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
        //    object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
        //    _wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
        //    ref nmatchAllWordForms, ref forward,
        //    ref wrap, ref format, ref replaceWithText,
        //    ref replaceAll, ref matchKashida,
        //    ref matchDiacritics, ref matchAlefHamza,
        //    ref matchControl);
        //}

        public void InitializeWordApp()
        {
            //MessageBox.Show("here");
            if (_wordApp == null) _wordApp = new ApplicationClass();

            deactivateevents = true;
            try
            {

                _wordApp.CommandBars.AdaptiveMenus = false;
                _wordApp.DocumentBeforePrint += new ApplicationEvents4_DocumentBeforePrintEventHandler(_wordApp_DocumentBeforePrint);
                _wordApp.DocumentBeforeSave += new ApplicationEvents4_DocumentBeforeSaveEventHandler(_wordApp_DocumentBeforeSave);
                //   _wordApp.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(OnOpenDoc);
                //  _wordApp.ApplicationEvents2_Event_Quit += new ApplicationEvents2_QuitEventHandler(OnQuit);
            }
            catch (Exception err) { MessageBox.Show(err.Message); }

            if (_hWnd == 0)
            {
                if (_wordDocumentCaption == null)
                {
                    _hWnd = FindWindow("Opusapp", null);
                }
                else
                {

                    _hWnd = FindWindow("Opusapp", _wordDocumentCaption + " - Microsoft Word");
                }
            }

            if (_hWnd != 0)
            {
                SetParent(_hWnd, this.Handle.ToInt32());


               
                /// We want to remove the system menu also. The title bar is not visible, but we want to avoid accidental minimize, maximize, etc ..by disabling the system menu(Alt+Space)
                try
                {
                    int hMenu = GetSystemMenu(_hWnd, false);
                    if (hMenu > 0)
                    {
                        int menuItemCount = GetMenuItemCount(hMenu);
                        DrawMenuBar(_hWnd);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                //geting all macro from server and add to microsoft
                //success = true;
                HttpWebResponse _webResponse = null;
                string url = App.ServerProtocol + App.ServerIP + ":" + Convert.ToString(App.ServerPort) + "/MSWORD/getAllMacro";
                HttpWebRequest _webRequest = (HttpWebRequest)WebRequest.Create(url);
                Cookie sessionCookie = new Cookie();
                sessionCookie.HttpOnly = true;
                sessionCookie.Domain = App.ServerIP;
                //TODO: see how port needs to be used.
                //sessionCookie.Port=App.ServerPort;
                sessionCookie.Name = "session_id";
                sessionCookie.Value = App.SessionId;
                _webRequest.CookieContainer = new CookieContainer();
                _webRequest.CookieContainer.Add(sessionCookie);
                _webRequest.Method = "GET";
                _webResponse = (HttpWebResponse)_webRequest.GetResponse();
                System.IO.StreamReader _streamReader = new System.IO.StreamReader(_webResponse.GetResponseStream());
                string jsonFormattedResponse = _streamReader.ReadToEnd();
                //logger.Debug("The response from the macro operation : " + jsonFormattedResponse);
                JSONStore.DeserialiseMacro(jsonFormattedResponse);
                //JSONStore.Macros.macro;
                //for each in 
                Dictionaries dict;
                //for (var i = 0; i <= len(JSONStore.Macros.macro); i++)
                //{
                //    //dict = JSONStore.Macros.macro[i];
                //}
                for (var i = 0; i < JSONStore.Macros.macro.Count; i++)
                {
                    _wordApp.Application.AutoCorrect.Entries.Add(JSONStore.Macros.macro[i].key, JSONStore.Macros.macro[i].value);
                }
                if (jsonFormattedResponse == "false")
                {
                    //logger.Info("Failed to Save Report in Data Base");
                    //success = false;
                }
                else
                {
                    //logger.Info("success to get Macro from Data Base");
                }
                //JSONStore.GetDataObject;
                //logger.Info("The get Macro activity completed");
                _wordApp.Application.AutoCorrect.Entries.Add("r", "rishu");
            }

        }

        void _wordApp_DocumentBeforeSave(Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Cancel = true;
        }

        void _wordApp_DocumentBeforePrint(Document Doc, ref bool Cancel)
        {
            Cancel = true;
        }

        private void acitvateWordApp()
        {
            try
            {
                _wordApp.Visible = true;
                _wordApp.Activate();
                SetWindowPos(_hWnd, this.Handle.ToInt32(), 0, 0, this.Bounds.Width, this.Bounds.Height, SWP_NOZORDER | SWP_NOMOVE | SWP_DRAWFRAME | SWP_NOSIZE);
                //Call onresize--I dont want to write the same lines twice
                OnResize();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: do not load the document into the control until the parent window is shown! {0}",ex.ToString());
            }

        }

        private void getBookmarks()
        {
            foreach (Bookmark bookmark in _wordDocument.Bookmarks)
            {
                // Find the report and Impression bookmarks.
                switch(bookmark.Name.ToUpper())
                {
                    case "REPORT_TEXT":
                        {
                            this._userReportBookmark = bookmark;
                        }
                        break;
                    case "IMPRESSION_TEXT":
                        {
                            this._userImpressionBookmark = bookmark;
                        }
                        break;
                    case "ADDENDUM_TEXT":
                        {
                            this._userAddendumText = bookmark;
                        }
                        break;
                    case "ADDENDUM_TITLE":
                        {
                            this._userAddendumTitle = bookmark;
                        }
                        break;

                }
            }
        }

        private void InitializeActiveWindow()
        {

            try
            {
                

                _wordApp.ActiveWindow.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                _wordApp.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = 0.0F;
                _wordDocumentCaption = Convert.ToString(Process.GetCurrentProcess().Id);
                _wordApp.ActiveWindow.DisplayRightRuler = false;
                _wordApp.ActiveWindow.DisplayScreenTips = false;
                _wordApp.ActiveWindow.DisplayVerticalRuler = false;
                _wordApp.ActiveWindow.DisplayRightRuler = false;
                _wordApp.ActiveWindow.ActivePane.DisplayRulers = false;
                bool temp=_wordApp.ActiveWindow.Document.ReadOnly;

                //_wordApp.ActiveWindow.Application.CommandBars.ActiveMenuBar.Visible = false;
                //_wordApp.ActiveWindow.Application.CommandBars = false;
                // _wordApp.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdWebView;
                //_wordApp.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;//wdWebView; // .wdNormalView;

                

                int counter = _wordApp.ActiveWindow.Application.CommandBars.Count;
                for (int i = 1; i <= counter; i++)
                {
                    try
                    {

                        String nm = _wordApp.ActiveWindow.Application.CommandBars[i].Name;
                        Microsoft.Office.Core.CommandBar cb = _wordApp.ActiveWindow.Application.CommandBars[i];
                        //if (nm == "Standard")
                        //{
                        //    //nm=i.ToString()+" "+nm;
                        //    //MessageBox.Show(nm);
                        //    int count_control = _wordApp.ActiveWindow.Application.CommandBars[i].Controls.Count;
                        //    for (int j = 1; j <= 3; j++)
                        //    {
                        //       //MessageBox.Show(_wordApp.ActiveWindow.Application.CommandBars[i].Controls[j].ToString());
                        //        _wordApp.ActiveWindow.Application.CommandBars[i].Controls[j].Enabled = false;

                        //    }
                        //}
                

                        if (nm == "Menu Bar")
                        {
                            //To disable the menubar, use the following (1) line
                            _wordApp.ActiveWindow.Application.CommandBars[i].Enabled = false;
                        }
                    }
                    catch
                    {

                    }
                }
                //Microsoft.Office.Core.CommandBar cb2=_wordApp.CommandBars["File"];

                //for (int i = 1; i <=cb2.Controls.Count; ++i)
                //{
                //    cb2.Controls[i].Enabled = false;
                //}



            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Contact The Server Please Try Agrin !!");
            }
        }

        private void InitializeActiveDocument()
        {

        }

        public void LoadBlankTemplate(bool InitHeader)
        {
            if (InitHeader == null)
                InitHeader = false;
            deactivateevents = true;

            //if (_wordApp == null) _wordApp = new ApplicationClass();
            //try
            //{

            //    _wordApp.CommandBars.AdaptiveMenus = false;
            // //   _wordApp.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(OnOpenDoc);
            //  //  _wordApp.ApplicationEvents2_Event_Quit += new ApplicationEvents2_QuitEventHandler(OnQuit);
            //}
            //catch (Exception err) { MessageBox.Show(err.Message); }
            if (_wordDocument != null)
            {
                try
                {
                    object dummy = null;
                    _wordApp.Documents.Close(ref dummy, ref dummy, ref dummy);
                }
                catch (Exception err)
                {
                    //     MessageBox.Show(err.Message);
                }
            }

            //if (_hWnd == 0)
            //{
            //    if (_wordDocumentCaption == null)
            //    {
            //        _hWnd = FindWindow("Opusapp", null);
            //    }
            //    else
            //    {

            //        _hWnd = FindWindow("Opusapp", _wordDocumentCaption + " - Microsoft Word");
            //    }
            //}
            if (_hWnd != 0)
            {
              //  SetParent(_hWnd, this.Handle.ToInt32());
                

                object fileName = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".docx";
                File.Copy(_blankTemplatePath, fileName.ToString());
                object newTemplate = false;
                object docType = 0;
                object readOnly = true;
                object isVisible = false;
                object missing = System.Reflection.Missing.Value;
                try
                {
                    if (_wordApp == null)
                    {
                        throw new WordInstanceException();
                    }

                    if (_wordApp.Documents == null)
                    {
                        throw new DocumentInstanceException();
                    }

                    if (_wordApp != null && _wordApp.Documents != null)
                    {
                        _wordApp.Visible = false;
                        _wordDocument = _wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing,
                            ref missing,ref missing,ref missing,ref missing,ref missing,ref missing,ref missing,
                            ref missing,ref missing,ref missing,ref missing,ref missing);
                        
                        InitializeActiveWindow();
                        //_wordApp.ActiveWindow.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                        //_wordApp.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = 0.0F;
                        //_wordDocumentCaption = Convert.ToString(Process.GetCurrentProcess().Id);
                       // //_wordApp.ActiveWindow.Caption = _wordDocumentCaption;
                        // try setting font.
                        //_wordDocument.Paragraphs[1].Range.Font.Name = "Times New Roman";
                        //_wordDocument.Paragraphs[1].Range.Font.Size = 12F;

                        // Page margin
                        //Note: DONT SET TOP MARGIN HERE.
                        // IT IS TAKEN CARE IN method: AddHeader
                        _wordDocument.PageSetup.BottomMargin = _wordApp.InchesToPoints((float)Convert.ToDouble(JSONStore.GetDataObject.bottomMargin));
                        _wordDocument.PageSetup.LeftMargin = _wordApp.InchesToPoints((float)Convert.ToDouble(JSONStore.GetDataObject.leftMargin));
                        _wordDocument.PageSetup.RightMargin = _wordApp.InchesToPoints((float)Convert.ToDouble(JSONStore.GetDataObject.rightMargin));

                        //if (InitHeader)
                        AddHeader1();
                       
                        ////    AddHeader(true);
                        //InitAndUpdateMergeFields();
                        MoveCursorToEndOfDocument();
                        this.blankTemplateWordCount = _wordDocument.Words.Count;
                        SetEditableArea();

                    }

                    if (_wordDocument == null)
                    {
                        throw new ValidDocumentException();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Unable to Contact The Server Please Try Agrin !!");
                }


                //try
                //{
                //    _wordApp.ActiveWindow.DisplayRightRuler = false;
                //    _wordApp.ActiveWindow.DisplayScreenTips = false;
                //    _wordApp.ActiveWindow.DisplayVerticalRuler = false;
                //    _wordApp.ActiveWindow.DisplayRightRuler = false;
                //    _wordApp.ActiveWindow.ActivePane.DisplayRulers = false;
                //    //_wordApp.ActiveWindow.Application.CommandBars.ActiveMenuBar.Visible = false;
                //    //_wordApp.ActiveWindow.Application.CommandBars = false;
                //    // _wordApp.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdWebView;
                //    //_wordApp.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;//wdWebView; // .wdNormalView;


                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("Unable to Contact The Server Please Try Agrin !!");
                //}



                //int counter = _wordApp.ActiveWindow.Application.CommandBars.Count;
                //for (int i = 1; i <= counter; i++)
                //{
                //    try
                //    {

                //        String nm = _wordApp.ActiveWindow.Application.CommandBars[i].Name;
                //        //        if (nm == "Standard")
                //        //        {
                //        //            //nm=i.ToString()+" "+nm;
                //        //            //MessageBox.Show(nm);
                //        //            int count_control = wd.ActiveWindow.Application.CommandBars[i].Controls.Count;
                //        //            for (int j = 1; j <= 3; j++)
                //        //            {
                //        //                //MessageBox.Show(_wordApp.ActiveWindow.Application.CommandBars[i].Controls[j].ToString());
                //        //                _wordApp.ActiveWindow.Application.CommandBars[i].Controls[j].Enabled = false;

                //        //            }
                //        //        }

                //        if (nm == "Menu Bar")
                //        {
                //            //To disable the menubar, use the following (1) line
                //            _wordApp.ActiveWindow.Application.CommandBars[i].Enabled = false;
                //        }
                //    }
                //    catch
                //    {

                //    }
                //}

                //            /// If you want to have specific menu or sub-menu items, write the code here. 
                //            /// Samples commented below
                //            /// 


                //            //							MessageBox.Show(nm);
                //            int count_control = _wordApp.ActiveWindow.Application.CommandBars[i].Controls.Count;
                //            //MessageBox.Show(count_control.ToString());						

                //            for (int j = 1; j <= count_control; j++)
                //            {
                //                /// The following can be used to disable specific menuitems in the menubar	
                //                _wordApp.ActiveWindow.Application.CommandBars[i].Controls[j].Enabled = true;

                //                ///The following can be used to disable some or all the sub-menuitems in the menubar


                //                Office.CommandBarPopup c;
                //                c = (Office.CommandBarPopup)_wordApp.ActiveWindow.Application.CommandBars[i].Controls[1];

                //                for (int k = 1; k <= c.Controls.Count; k++)
                //                {
                //                    //MessageBox.Show(k.ToString()+" "+c.Controls[k].Caption + " -- " + c.Controls[k].DescriptionText + " -- " );
                //                    try
                //                    {
                //                        c.Controls[1].Enabled = false;
                //                        c.Controls[2].Enabled = false;
                //                        c.Controls[3].Enabled = false;
                //                        c.Controls[4].Enabled = false;
                //                        c.Controls[5].Enabled = false;
                //                        c.Controls["Close Window"].Enabled = false;
                //                    }
                //                    catch
                //                    {

                //                    }
                //                }
                //                //_wordApp.ActiveWindow.Application.CommandBars[i].Controls[j].Control	 Controls[0].Enabled=false;
                //            }
                //        }

                //        nm = "";
                //    }
                //    catch (Exception ex)
                //    {
                //        MessageBox.Show(ex.ToString());
                //    }
                //}
                // Show the word-document
                //try
                //{
                //    _wordApp.Visible = true;
                //    _wordApp.Activate();
                //    SetWindowPos(_hWnd, this.Handle.ToInt32(), 0, 0, this.Bounds.Width, this.Bounds.Height, SWP_NOZORDER | SWP_NOMOVE | SWP_DRAWFRAME | SWP_NOSIZE);
                //    //Call onresize--I dont want to write the same lines twice
                //    OnResize();
                //}
                //catch( Exception ex)
                //{
                //    MessageBox.Show("Error: do not load the document into the control until the parent window is shown!");
                //}

                ///// We want to remove the system menu also. The title bar is not visible, but we want to avoid accidental minimize, maximize, etc ..by disabling the system menu(Alt+Space)
                //try
                //{
                //    int hMenu = GetSystemMenu(_hWnd, false);
                //    if (hMenu > 0)
                //    {
                //        int menuItemCount = GetMenuItemCount(hMenu);
                //        DrawMenuBar(_hWnd);
                //    }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message);
                //}
                // this.Parent.Focus();

                acitvateWordApp();
                
            }
            deactivateevents = false;
        }

        private void SetEditableArea()
        {
            
            getBookmarks();

            // set the access
            _wordDocument.Protect(WdProtectionType.wdAllowOnlyReading, ref _noReset, ref _password, ref _useIRM, ref _enforceStyleLock);
            object EditAccess = WdEditorType.wdEditorEveryone;
            object NoAccess = WdEditorType.wdEditorOwners;
            if (JSONStore.GetDataObject.status.ToUpper() != "SIGNEDOFF")
            {
                try
                {
                    _userReportBookmark.Range.Editors.Add(ref EditAccess);
                    _userImpressionBookmark.Range.Editors.Add(ref EditAccess);
                }
                catch (Exception)
                { }
            }
            //else{
            //    UnProtectFile();
            //    _wordDocument.Protect(WdProtectionType.wdAllowOnlyReading, ref _noReset, ref _password, ref _useIRM, ref _enforceStyleLock);



            //}
        }
        
        public void setAddendumTextArea()
        {


            
            
            // set the access
            getBookmarks();

            if (_wordDocument.ProtectionType == WdProtectionType.wdNoProtection)
                _wordDocument.Protect(WdProtectionType.wdAllowOnlyReading, ref _noReset, ref _password, ref _useIRM, ref _enforceStyleLock);
            //else
            //{
            //    _wordDocument.Protect(
            //}
            object EditAccess = WdEditorType.wdEditorEveryone;
            object NoAccess = WdEditorType.wdEditorOwners;

            //_userReportBookmark.

            bool isempty = false;

            if (_userAddendumTitle.Empty)
                isempty = true;

            try
            {

                if (_userAddendumTitle.Range.Text.Trim().Length == 0)
                    isempty = true;
            }
            catch (Exception)
            { }

            if(isempty)    // do only if it has not been yet added
            {
                string radiologistName = JSONStore.GetDataObject.userData.fname + " " + JSONStore.GetDataObject.userData.mname + " " + JSONStore.GetDataObject.userData.lname;
                string reportedDateAndTime = Utils.MyDateTimeFormat.GetCurrentDateTimeFormattedForReport();
                string addendumText = Environment.NewLine + Environment.NewLine + "ADDENDUM (By: " + radiologistName + " " + " Date: " + reportedDateAndTime + ")";

                Clipboard.Clear();
                Clipboard.SetText(addendumText, TextDataFormat.Rtf);
                _userAddendumText.Range.Editors.Add(ref EditAccess);


                
                _userAddendumTitle.Range.Editors.Add(ref EditAccess);
              

                _userAddendumTitle.Select();

                object rangeBookmark = _wordApp.Selection.Range;                
                _wordApp.Selection.Paste();
                //_wordApp.Selection.Paste();

                _wordApp.Selection.Range.Text = string.Empty;
                _wordApp.Selection.Range.Underline = WdUnderline.wdUnderlineDash;

                if (!_wordDocument.Bookmarks.Exists("ADDENDUM_TITLE"))
                {
                    _wordDocument.Bookmarks.Add("ADDENDUM_TITLE", ref rangeBookmark);
                }

                
                //_userAddendumTitle.Range.Font = Microsoft.Office.Interop.Word.Font Font.Bold;
                _userAddendumTitle.Range.Editors.Add(ref NoAccess);

                object bookmarkName = "ADDENDUM_TEXT";
                if (_wordDocument.Bookmarks.Exists(bookmarkName.ToString()))
                {
                    Bookmark bookmark = _wordDocument.Bookmarks.get_Item(ref bookmarkName);
                    bookmark.Select();
                }
            }
        }
        public void OnResize()
        {
            //int borderWidth = SystemInformation.Border3DSize.Width;
            //int borderHeight = SystemInformation.Border3DSize.Height;
            //int captionHeight = SystemInformation.CaptionHeight;
            //int statusHeight = SystemInformation.ToolWindowCaptionHeight;
            //MoveWindow(
            //    _hWnd,
            //    -2 * borderWidth,
            //    -2 * borderHeight - captionHeight,
            //    Convert.ToInt32(this.Width + 4 * borderWidth),
            //    Convert.ToInt32(this.Height + captionHeight + 4 * borderHeight + statusHeight),
            //    true);



            //
            // Due to containment relationship changed. the above is not valid.
            // So using the docking center code.
            MoveWindow(_hWnd, 0, 0, this.Width, this.Height,true);
        }

        private void OnResize(object sender, System.EventArgs e)
        {
            OnResize();
        }


        public void RestoreCommandBars()
        {
            try
            {
                int counter = _wordApp.ActiveWindow.Application.CommandBars.Count;
                for (int i = 1; i <= counter; i++)
                {
                    try
                    {

                        String nm = _wordApp.ActiveWindow.Application.CommandBars[i].Name;
                        if (nm == "Standard")
                        {
                            int count_control = _wordApp.ActiveWindow.Application.CommandBars[i].Controls.Count;
                            for (int j = 1; j <= 3; j++)
                            {
                                _wordApp.ActiveWindow.Application.CommandBars[i].Controls[j].Enabled = false;
                            }
                        }
                        if (nm == "Menu Bar")
                        {
                            _wordApp.ActiveWindow.Application.CommandBars[i].Enabled = false;
                        }
                        nm = "";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public class DocumentInstanceException : Exception
        { }

        public class ValidDocumentException : Exception
        { }

        public class WordInstanceException : Exception
        { }


        internal void LoadTemplate(Template template)
        {
            if (template.content == string.Empty)
            {

            }
            else
            {
                //<maisee> first unlock the document 
                //this.UnProtectFile();
                Clipboard.Clear();
                Clipboard.SetText(template.content, TextDataFormat.Rtf);

                getBookmarks();
                _userReportBookmark.Select();
                _userReportBookmark.Range.Text = string.Empty;
                object rangeBookmark = _wordApp.Selection.Range;
                _wordApp.Selection.Paste();

                if (!_wordDocument.Bookmarks.Exists("REPORT_TEXT"))
                {
                    _wordDocument.Bookmarks.Add("REPORT_TEXT", ref rangeBookmark);
                }

                //<maisee> lock the doc again
                //SetEditableArea();

            }
            //LoadReportingDoctorNameAndSignature();
         //   MoveCursorToStartOfDocument();
           // _wordDocument.Paragraphs[1].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

        }

        internal void LoadReportingDoctorNameAndSignature(string singatureBits)
        {
            MoveCursorToEndOfDocument();

            _wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            _wordApp.Selection.ParagraphFormat.SpaceAfter = 0.0F;
            _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            _wordApp.Selection.InsertParagraphAfter();
            _wordApp.Selection.TypeParagraph();
            WdColor oldColor = _wordApp.Selection.Font.Color;
            _wordApp.Selection.Font.Color = WdColor.wdColorWhite;            
            _wordApp.Selection.TypeText("__________________________________________________________________________________________________");
            _wordApp.Selection.TypeText(Environment.NewLine);
            _wordApp.Selection.Font.Color = oldColor;
            //object row = 1;
            //_wordApp.Selection.InsertRows(ref row);
            if (String.Compare(singatureBits, "No signature found", true) == 0)
            {

            }
            else
            {
                
                
                Clipboard.Clear();
                int discardedBytes = 0;
                byte[] signature = Utils.GetBytes(singatureBits, out discardedBytes);
               

                ImageConverter ic = new ImageConverter();
                Image img = (Image)ic.ConvertFrom(signature);
                Clipboard.SetImage(img);
                _wordApp.Selection.Paste();
                _wordApp.Selection.TypeText(Environment.NewLine) ;
            }
            //_wordApp.Selection.TypeParagraph();           

            var userName = "";
            var degree = "";
            var desg = "";
            if (JSONStore.GetDataObject.userData.morphedRadiologist != "" && JSONStore.GetDataObject.userData.morphedRadiologist != null)
            {
                userName = JSONStore.GetDataObject.userData.morphedRadFname + " " + JSONStore.GetDataObject.userData.morphedRadMname + " " + JSONStore.GetDataObject.userData.morphedRadLname;
                degree = JSONStore.GetDataObject.userData.morphedRadDegree;
                desg = JSONStore.GetDataObject.userData.morphedRadDesg;
            }
            else
            {
                userName = JSONStore.GetDataObject.userData.fname + " " + JSONStore.GetDataObject.userData.mname + " " + JSONStore.GetDataObject.userData.lname;
                degree = JSONStore.GetDataObject.userData.degree;
                desg = JSONStore.GetDataObject.userData.desg;
            }

            _wordApp.Selection.TypeText(userName);
            _wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
       
           
            if ((bool)JSONStore.GetDataObject.showDoctorDesgBeforeDegree == true)
            {
                if (desg.Length > 0)
                {
                    _wordApp.Selection.TypeParagraph();
                    _wordApp.Selection.TypeText(desg);
                    
                }
                if (degree.Trim().Length > 0)
                {
                    _wordApp.Selection.TypeParagraph();
                    _wordApp.Selection.TypeText(degree);
                   
                }
            }
            else 
            {
                if (degree.Trim().Length > 0)
                {
                    _wordApp.Selection.TypeParagraph();
                    _wordApp.Selection.TypeText(JSONStore.GetDataObject.userData.degree);
                    
                }
                if (desg.Trim().Length > 0)
                {
                    _wordApp.Selection.TypeParagraph();
                    _wordApp.Selection.TypeText(JSONStore.GetDataObject.userData.desg);
                    
                }
                
            }
           
            //_wordApp.Selection.TypeText(JSONStore.GetDataObject.userData.signature);
            //if (String.Compare(JSONStore.GetDataObject.userData.signature, "No signature found", true) == 0)
            //{

            //}
            //else
            //{
            //    Clipboard.Clear();
            //    int discardedBytes = 0;
            //    byte[] signature = Utils.GetBytes(JSONStore.GetDataObject.userData.signature, out discardedBytes);
            //    ImageConverter ic = new ImageConverter();
            //    Image img = (Image)ic.ConvertFrom(signature);


            //    Clipboard.SetImage(img);
            //    _wordApp.Selection.Paste();

            //}
            MoveCursorToStartOfDocument();
            _wordDocument.Paragraphs[1].Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
          
        }


        internal void findAndReplaceTextInHeaderFooter(string findText, string replaceText, bool wildCard, WdColor wdColor)
        {
            //logger.Debug(" Asked me to find :" + findText + " ====> replace text : " + replaceText + " wildCard : " + wildCard + " WdColor : " + wdColor);

            //logger.Debug("Word.Document : " + _wordDocument.Sections.Count.ToString());
            foreach (Section section in _wordDocument.Sections)
            {
                for (int i = 1; i < 3; i++)
                {
                    Range _rng = section.Footers[(WdHeaderFooterIndex)i].Range;
                    //object EditAccess = WdEditorType.wdEditorEveryone;

                    //_rng.Editors.Add(ref EditAccess);
                    //  MessageBox.Show(_rng.Text);
                    //  _rng.Find.MatchWildcards = true;
                    //this.FindAndReplaceAll(ref _rng,regex, currRepDateTime);
                    object replaceAll = WdReplace.wdReplaceAll;
                    object missing = System.Type.Missing;
                    _rng.Find.Text = findText;
                    _rng.Find.Replacement.Text = replaceText;
                    _rng.Find.MatchWildcards = wildCard;
                    if (wdColor != WdColor.wdColorAutomatic)
                        _rng.Find.Replacement.Font.Color = wdColor;
                    // Execute the Find and Replace -- notice that the
                    // 11th parameter is the "replaceAll" enum object
                    bool res = _rng.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);
                    // MessageBox.Show(_rng.Text);
                    _rng.Fields.Update();


                    _rng = section.Headers[(WdHeaderFooterIndex)i].Range;
                    //_rng.Editors.Add(ref EditAccess);
                    _rng.Find.Text = findText;
                    _rng.Find.Replacement.Text = replaceText;
                    if (wdColor != WdColor.wdColorAutomatic)
                        _rng.Find.Replacement.Font.Color = wdColor;

                    _rng.Find.MatchWildcards = wildCard;
                    res = _rng.Find.Execute(ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref replaceAll,
                        ref missing, ref missing, ref missing, ref missing);

                    _rng.Fields.Update();

                }
            }
            // update fields in document.
            //_wordDocument.Fields.Update();

        }
        internal void SaveDocument(string destFile)
        {
            this.ServerDepReportDateTime = App.GetServerDpeReportDateTime();           
            this.reportedDateAndTime = this.ServerDepReportDateTime.ToString(Utils.MyDateTimeFormat.FmtDateTime);         

            string currRepDateTime = this.reportedDateAndTime + JSONStore.GetDataObject.textInvisibleMarker.ToString();
            //PreviousReport pr = this.AdditionalInformation.getPreviousReport(JSONStore.GetDataObject.latestReportId);

            // Update in header and footer
            //foreach (Section section in _wordDocument.Sections)
            //{
            //    Range _rng = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            //    this.FindAndReplaceAll(ref _rng, this.reportedDateAndTime,currRepDateTime);
            //    _rng.Fields.Update();
                
            //    _rng = section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
            //     this.FindAndReplaceAll(ref _rng, this.reportedDateAndTime,currRepDateTime);
            //    _rng.Fields.Update();
            //}

            //string regex = "REPORT TIME: #[0-9]{2}-[0-9]{2}-[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2}#";
            //string regex = "#([^a-zA-Z])*#";
            string regex = "[ :0-9-]{19}" + JSONStore.GetDataObject.textInvisibleMarker.ToString();
            //string regex = "REPORT TIME";
            this.findAndReplaceTextInHeaderFooter(regex, currRepDateTime, true, WdColor.wdColorAutomatic);
            this.findAndReplaceTextInHeaderFooter(JSONStore.GetDataObject.textInvisibleMarker.ToString(), JSONStore.GetDataObject.textInvisibleMarker.ToString(), false, WdColor.wdColorWhite);
            //foreach (Section section in _wordDocument.Sections)
            //{
            //    for (int i = 1; i < 3; i++)
            //    {
            //        Range _rng = section.Footers[(WdHeaderFooterIndex)i].Range;
            //        //  MessageBox.Show(_rng.Text);
            //        //  _rng.Find.MatchWildcards = true;
            //        //this.FindAndReplaceAll(ref _rng,regex, currRepDateTime);
            //        object replaceAll = WdReplace.wdReplaceAll;
            //        object missing = System.Type.Missing;
            //        _rng.Find.Text = regex;
            //        _rng.Find.Replacement.Text = currRepDateTime;
            //        _rng.Find.MatchWildcards = true;
            //        // Execute the Find and Replace -- notice that the
            //        // 11th parameter is the "replaceAll" enum object
            //        bool res = _rng.Find.Execute(ref missing, ref missing, ref missing,
            //            ref missing, ref missing, ref missing, ref missing,
            //            ref missing, ref missing, ref missing, ref replaceAll,
            //            ref missing, ref missing, ref missing, ref missing);
            //        // MessageBox.Show(_rng.Text);
            //        _rng.Fields.Update();

                
            //        _rng = section.Headers[(WdHeaderFooterIndex)i].Range;
            //        _rng.Find.Text = regex;
            //        _rng.Find.Replacement.Text = currRepDateTime;
            //        _rng.Find.MatchWildcards = true;
            //        res = _rng.Find.Execute(ref missing, ref missing, ref missing,
            //            ref missing, ref missing, ref missing, ref missing,
            //            ref missing, ref missing, ref missing, ref replaceAll,
            //            ref missing, ref missing, ref missing, ref missing);

            //        _rng.Fields.Update();
                
            //    }
            //}
            //// update fields in document.
            //_wordDocument.Fields.Update();

            object _unknown = Type.Missing;
            object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;
            _wordDocument.Save();

            File.Copy(_wordDocument.FullName, destFile);
            //_wordApp.ActiveDocument.SaveAs(ref destFile, ref format, ref _unknown, ref _unknown, ref _unknown,
            //    ref _unknown, ref _unknown, ref _unknown, ref _unknown, ref _unknown, ref _unknown, ref _unknown,
            //    ref _unknown, ref _unknown, ref _unknown, ref _unknown);
            
            //wordApp.ActiveWindow.Caption = _wordDocumentCaption;
            ////_wordApp.ActiveWindow.Caption = _wordDocumentCaption;
            //close the document after saving.
            //caller needs to handle for opening blank document again.
            //CloseDocument();
        }

        internal void LoadReport()
        {

            int discardedBytes;
            byte[] reportData = Utils.GetBytes(JSONStore.GetDataObject.latestReportData, out discardedBytes);
            object fileName = Utils.GetLocalDir + "\\" + JSONStore.GetDataObject.latestReportId + ".docx";


            FileStream fs = new FileStream((string)fileName, FileMode.OpenOrCreate);
            fs.Write(reportData, 0, reportData.Length);
            fs.Close();
            CloseDocument();

            object newTemplate = false;
            object docType = 0;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;

            _wordDocument = _wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing
                , ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //_wordApp.ActiveWindow.Caption = _wordDocumentCaption;
            //_wordApp.ActiveWindow.Caption = _wordDocumentCaption;
            MoveCursorToStartOfDocument();
            _wordApp.Visible = true;
            this.acitvateWordApp();

            SetEditableArea();
            
        }
        internal void OpenFile(string p)
        {
            object fileName = p;
            object newTemplate = false;
            object docType = 0;
            object readOnly = true;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;
            _wordDocument = _wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing
                , ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
//            _wordDocument = _wordApp.Documents.Add(ref fileName, ref newTemplate, ref docType, ref isVisible);
            //_wordApp.ActiveWindow.Caption = _wordDocumentCaption;
            //_wordApp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100;
            SetEditableArea();
        }



        internal bool IsBlankTemplate()
        {

            if (this._wordDocument.Words.Count > this.blankTemplateWordCount)
            {
                return false;
            }
            return true;
        }

        internal void UnProtectFile()
        {
            try
            {
                _wordDocument.Unprotect(ref _password); 
            }
            catch(Exception)
            {
            }
        }

        internal void UpdateAddendumBookmarks()
        {
            getBookmarks();
            try
            {
                if (_userAddendumText != null)
                {
                    if (_userAddendumText.Range.Text.Trim().Length == 0)
                        return;
                    // old addendum found .. add a new book mark 

                    Clipboard.Clear();
                    Clipboard.SetText(Environment.NewLine + " ", TextDataFormat.Rtf);

                    _userAddendumText.Select();
                    _wordApp.Selection.InsertAfter(Environment.NewLine + " ");

                    object NewEndPos = _wordApp.Selection.Range.StoryLength - 1;
                    Range rng = _wordDocument.Range(ref NewEndPos, ref NewEndPos);

                    rng.Select();

                    object rangeBookmark = _wordApp.Selection.Range;
                    getBookmarks();

                    if (_userAddendumTitle != null)
                    {
                        _userAddendumTitle.Delete();
                    }

                    _wordApp.Selection.InsertAfter(Environment.NewLine + Environment.NewLine + " ");

                    _wordDocument.Bookmarks.Add("ADDENDUM_TITLE", ref rangeBookmark);

                    NewEndPos = _wordApp.Selection.StoryLength - 1;
                    rng = _wordDocument.Range(ref NewEndPos, ref NewEndPos);

                    rng.Select();

                    rangeBookmark = _wordApp.Selection.Range;

                    getBookmarks();

                    if (_userAddendumText != null)
                    {
                        try
                        {
                            _userAddendumText.Delete();
                        }
                        catch (Exception)
                        {
                        }
                    }
                    _wordDocument.Bookmarks.Add("ADDENDUM_TEXT", ref rangeBookmark);

                
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
