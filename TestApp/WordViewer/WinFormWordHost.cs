#region includes
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.Word;
using System.IO;
#endregion
namespace Meddiff.Common.WordViewer
{
    public partial class WinFormWordHost : UserControl
    {
       
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

        #region "Properties and varibles"
        //Represent word document
        public Document _wordDocument;

        private ApplicationClass _wordApp = null;
        //Represent word application
        public ApplicationClass WordApp
        {
            get{ return _wordApp; }
        }

        //Represent application IntPtr handle.
        private static int _hWnd = 0;
        
        public string InputFile
        {
            get { return InputFile; }
        }
        
        private string _wordDocumentCaption = string.Empty;

        //Represent a file that will be loaded.
        string _inpFilePath = string.Empty;

        private string _blankTemplatePath = @"E:\Make Web Compatible to IE 11.docx";
        /// <summary>
        /// Blank template path 
        /// </summary>
        public string BlankTemplatePath
        {
            get { return _blankTemplatePath; }
        }

        #endregion "Properties and varibles"

        public WinFormWordHost()
        {
            InitializeComponent();
            this.InitializeWordApp();


        }

        private void InitializeWordApp()
        {
            if (_wordApp == null) _wordApp = new ApplicationClass();
            if (_hWnd == 0)
            {
                if (string.IsNullOrEmpty(_wordDocumentCaption))
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

                 LoadBlankTemplate();
                 //acitvateWordApp();
             }
        }
        private void acitvateWordApp()
        {
            try
            {
                _wordApp.Visible = true;
                _wordApp.Activate();
               
            }
            catch (Exception)
            {
               // MessageBox.Show("Error: do not load the document into the control until the parent window is shown! {0}", ex.ToString());
            }

        }

        public void LoadBlankTemplate(bool InitHeader = false)
        {
            //if (_wordDocument != null)
            //{
            //    try
            //    {
            //        //object dummy = null;
            //        //_wordApp.Documents.Close(ref dummy, ref dummy, ref dummy);
            //    }
            //    catch (Exception err)
            //    {
            //            MessageBox.Show(err.Message);
            //    }
            //}
            if (_hWnd != 0)
            {
                object fileName = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".docx";
                File.Copy(_blankTemplatePath, fileName.ToString()); // create a temp file  
                object newTemplate = false;
                object docType = 0;
                object readOnly = true;
                object isVisible = false;
                object missing = System.Reflection.Missing.Value;
                try
                {
                    if (_wordApp == null)
                    {
                        //throw new WordInstanceException();
                    }

                    if (_wordApp.Documents == null)
                    {
                        //throw new DocumentInstanceException();
                    }

                    if (_wordApp != null && _wordApp.Documents != null)
                    {
                        _wordApp.Visible = false;
                        _wordDocument = _wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing);

                    }

                    if (_wordDocument == null)
                    {
                        //throw new ValidDocumentException();
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Unable to Contact The Server Please Try Agrin !!");
                }

                acitvateWordApp();

            }
            //deactivateevents = false;
        }

        internal void SetShowToolBar(bool _showToolBar)
        {
            //_wordApp.
        }
    }
}
