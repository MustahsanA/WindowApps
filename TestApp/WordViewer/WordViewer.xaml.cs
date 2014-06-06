using Microsoft.Office.Interop.Word;
using System;
using System.Windows.Controls;

namespace Meddiff.Common.WordViewer
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class WordViewer : UserControl
    {

        #region "Password Protect"
        object _noReset = false;
        object _password = "MeddiffTech";
        object _useIRM = false;
        object _enforceStyleLock = false;
        #endregion

        Document _wordDocument;
        private readonly MSWordHost _winFormWordHost;

        public WordViewer()
        {
            InitializeComponent();
            _winFormWordHost = wordHost.Child as MSWordHost;
        }


        string _wordPath;
        public string WordPath
        {
            get { return _wordPath; }
            set
            {
                _wordPath = value;
                LoadFile(_wordPath);
            }
        }

        


        public void LoadFile(string path)
        {
            try
            {
                _wordPath = path;

                // now load
                object missing = System.Reflection.Missing.Value;
                object fileName = path;
                _wordDocument = _winFormWordHost.WordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            catch(Exception)
            {
                //TODO: Add Log
            }
        }


        public void SaveFile(string path = "")
        {
            if ( string.IsNullOrEmpty(path))
            {
                //just save at the same location

                try
                {

                    if (_wordDocument != null)
                    {
                        _wordDocument.Save();
                    }

                    return;
                }
                catch(Exception)
                {
                    //TODO: Add Log
                }
            }

            // saving as 
            try
            {
                object _unknown = Type.Missing;
                object format = WdSaveFormat.wdFormatDocumentDefault;
                object DestFilePath = path;

                _wordDocument.SaveAs(ref DestFilePath, ref format, ref _unknown, ref _unknown, ref _unknown,
                    ref _unknown, ref _unknown, ref _unknown, ref _unknown, ref _unknown, ref _unknown, ref _unknown,
                    ref _unknown, ref _unknown, ref _unknown, ref _unknown);
            }
            catch (Exception)
            {
                //TODO: Add Log
            }
        }

        void ProtectFile()
        {
            try
            {
                _wordDocument.Protect(WdProtectionType.wdAllowOnlyReading, ref _noReset, ref _password, ref _useIRM, ref _enforceStyleLock);

            }
            catch(Exception ex)
            {

            }

        }

        void UnProtectFile()
        {

        }

        
    }
}
