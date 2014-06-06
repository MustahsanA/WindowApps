using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace TestApp
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    //public partial class MessageBox : Window
    //{
    //    public MessageBox()
    //    {
    //        InitializeComponent();
    //    }
    //}
    public class MessageBox : Window
    {
        private MessageBoxButton _buttons = MessageBoxButton.OK;
        private string _message;
        private string _title;
        private MessageBoxResult _result = MessageBoxResult.None;

        public MessageBox(string message, string title, MessageBoxButton buttons)
        {
            Title = title;
            Message = message;
            Buttons = buttons;
        }

        public string Title
        {
            get { return _title; }
            set
            {
                _title = value;
                //NotifyOfPropertyChange(() => Title);
            }
        }

        public bool IsNoButtonVisible
        {
            get { return _buttons == MessageBoxButton.YesNo || _buttons == MessageBoxButton.YesNoCancel; }
        }

        public bool IsYesButtonVisible
        {
            get { return _buttons == MessageBoxButton.YesNo || _buttons == MessageBoxButton.YesNoCancel; }
        }

        public bool IsCancelButtonVisible
        {
            get { return _buttons == MessageBoxButton.OKCancel || _buttons == MessageBoxButton.YesNoCancel; }
        }

        public bool IsOkButtonVisible
        {
            get { return _buttons == MessageBoxButton.OK || _buttons == MessageBoxButton.OKCancel; }
        }

        public string Message
        {
            get { return _message; }
            set
            {
                _message = value;
                //NotifyOfPropertyChange(() => Message);
            }
        }

        public MessageBoxButton Buttons
        {
            get { return _buttons; }
            set
            {
                _buttons = value;
                //NotifyOfPropertyChange(() => IsNoButtonVisible);
                //NotifyOfPropertyChange(() => IsYesButtonVisible);
                //NotifyOfPropertyChange(() => IsCancelButtonVisible);
                //NotifyOfPropertyChange(() => IsOkButtonVisible);
            }
        }

        public MessageBoxResult Result { get { return _result; } }

        public void No()
        {
            _result = MessageBoxResult.No;
           // TryClose(false);
        }

        public void Yes()
        {
            _result = MessageBoxResult.Yes;
            //TryClose(true);
        }

        public void Cancel()
        {
            _result = MessageBoxResult.Cancel;
            //TryClose(false);
        }

        public void Ok()
        {
            _result = MessageBoxResult.OK;
           // TryClose(true);
        }

        //public virtual void TryClose(bool? dialogResult = null)
        //{
        //    PlatformProvider.Current.GetViewCloseAction(this, Views.Values, dialogResult).OnUIThread();
        //}
    }
}
