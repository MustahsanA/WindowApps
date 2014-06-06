using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApp
{
    public enum RadMessageWindowType
    {
        /// <summary>
        /// Message Box window displayed as Modal window.
        /// </summary>
        Modal,
        /// <summary>
        /// Message Box window displayed as Non-Modal window.
        /// </summary>
        NonModal
    }
    public enum RadMessageBoxResult
    {
        /// <summary>
        /// None of the buttons clicked.
        /// </summary>
        None = 1,
        /// <summary>
        /// Only OK button is clicked.
        /// </summary>
        Ok = 2,
        /// <summary>
        /// Cancel button is clicked.
        /// </summary>
        Cancel = 4,
        /// <summary>
        /// Yes button is clicked.
        /// </summary>
        Yes = 8,

        /// <summary>
        /// No button is clicked.
        /// </summary>
        No = 16,
    }

    public enum RadMessageBoxButton
    {
        // Summary:
        //     The message box displays an OK button.
        OK = 0,
        //
        // Summary:
        //     The message box displays OK and Cancel buttons.
        OKCancel = 1,
        //
        // Summary:
        //     The message box displays Yes, No, and Cancel buttons.
        YesNoCancel = 3,
        //
        // Summary:
        //     The message box displays Yes and No buttons.
        YesNo = 4,
    }

    public enum RadMessageBoxIcon
    {
        // Summary:
        //     No icon is displayed.
        None,
        //
        // Summary:
        //     The message box contains a symbol consisting of white X in a circle with
        //     a red background.
        Error,
        //
        // Summary:
        //     The message box contains a symbol consisting of a white X in a circle with
        //     a red background.
        Hand,
        //
        // Summary:
        //     The message box contains a symbol consisting of white X in a circle with
        //     a red background.
        Stop,
        //
        // Summary:
        //     The message box contains a symbol consisting of a question mark in a circle.
        Question,
        //
        // Summary:
        //     The message box contains a symbol consisting of an exclamation point in a
        //     triangle with a yellow background.
        Exclamation,
        //
        // Summary:
        //     The message box contains a symbol consisting of an exclamation point in a
        //     triangle with a yellow background.
        Warning,
        //
        // Summary:
        //     The message box contains a symbol consisting of a lowercase letter i in a
        //     circle.
        Information,
        //
        // Summary:
        //     The message box contains a symbol consisting of a lowercase letter i in a
        //     circle.
        Asterisk,

    }

    /// <summary>
    /// A converter helper functions class from RadMessageBox framework.
    /// </summary>
    class RadMsgConvert
    {
        /// <summary>
        /// Converts MessageBoxResults to RadMessageBoxResult.
        /// </summary>
        /// <param name="_in"> MessageBoxResult </param>
        /// <returns>RadMessageBoxResult</returns>
        public static RadMessageBoxResult ToRadMessageBoxResult(System.Windows.MessageBoxResult _in)
        {
            switch (_in)
            {
                case System.Windows.MessageBoxResult.None:
                    return RadMessageBoxResult.None;

                case System.Windows.MessageBoxResult.OK:
                    return RadMessageBoxResult.Ok;

                case System.Windows.MessageBoxResult.Cancel:
                    return RadMessageBoxResult.Cancel;

                case System.Windows.MessageBoxResult.No:
                    return RadMessageBoxResult.No;

                case System.Windows.MessageBoxResult.Yes:
                    return RadMessageBoxResult.Yes;

            }

            return RadMessageBoxResult.None;
        }

        /// <summary>
        /// Converts RadMessageBoxIcon to MessageBoxImage
        /// </summary>
        /// <param name="icon"> input param to convert</param>
        /// <returns>MessageBoxImage</returns>
        public static System.Windows.MessageBoxImage ToMessageBoxImage(RadMessageBoxIcon icon)
        {

            switch (icon)
            {
                case RadMessageBoxIcon.Asterisk:
                    return System.Windows.MessageBoxImage.Asterisk;
                case RadMessageBoxIcon.Error:
                    return System.Windows.MessageBoxImage.Error;
                case RadMessageBoxIcon.Exclamation:
                    return System.Windows.MessageBoxImage.Exclamation;
                case RadMessageBoxIcon.Hand:
                    return System.Windows.MessageBoxImage.Hand;
                case RadMessageBoxIcon.Information:
                    return System.Windows.MessageBoxImage.Information;
                case RadMessageBoxIcon.None:
                    return System.Windows.MessageBoxImage.None;
                case RadMessageBoxIcon.Question:
                    return System.Windows.MessageBoxImage.Question;
                case RadMessageBoxIcon.Stop:
                    return System.Windows.MessageBoxImage.Stop;
                case RadMessageBoxIcon.Warning:
                    return System.Windows.MessageBoxImage.Warning;

            }

            return System.Windows.MessageBoxImage.None;
        }

        /// <summary>
        /// Converts RadMessageBoxButton to MessageBoxButton
        /// </summary>
        /// <param name="button"></param>
        /// <returns></returns>
        public static System.Windows.MessageBoxButton ToMessageBoxButton(RadMessageBoxButton button)
        {

            switch (button)
            {
                case RadMessageBoxButton.OK:
                    return System.Windows.MessageBoxButton.OK;

                case RadMessageBoxButton.OKCancel:
                    return System.Windows.MessageBoxButton.OKCancel;
                case RadMessageBoxButton.YesNo:
                    return System.Windows.MessageBoxButton.YesNo;
                case RadMessageBoxButton.YesNoCancel:
                    return System.Windows.MessageBoxButton.YesNoCancel;
            }

            return System.Windows.MessageBoxButton.OK;
        }
    }
    /// <summary>
    /// MessageBox display wrapper framework main class.
    /// </summary>
    public class RadMessageBox
    {
        private readonly static string caption = "InstaWork Station";

        //show Modal message.
        public static RadMessageBoxResult ShowDialog(string message)
        {

            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message,caption);

            return RadMsgConvert.ToRadMessageBoxResult(result);

        }

        //Show Non - modal message.
        public static RadMessageBoxResult Show(string message)
        {

            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption);

            return RadMsgConvert.ToRadMessageBoxResult(result);

        }
        

        public static RadMessageBoxResult Show(string message,string caption)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption);
            return RadMsgConvert.ToRadMessageBoxResult(result);
        }


        public static RadMessageBoxResult ShowDialog(string message, string caption)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption);
            return RadMsgConvert.ToRadMessageBoxResult(result);
        }


        public static RadMessageBoxResult Show(string message, RadMessageBoxButton button)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption, RadMsgConvert.ToMessageBoxButton(button));
            return RadMsgConvert.ToRadMessageBoxResult(result);
        }

        public static RadMessageBoxResult ShowDialog(string message, RadMessageBoxButton button)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption, RadMsgConvert.ToMessageBoxButton(button));
            return RadMsgConvert.ToRadMessageBoxResult(result);
        }



        public static RadMessageBoxResult Show(string message, RadMessageBoxButton button, string caption )
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption, RadMsgConvert.ToMessageBoxButton(button));
            return RadMsgConvert.ToRadMessageBoxResult(result);
        }


        public static RadMessageBoxResult ShowDialog(string message, RadMessageBoxButton button, string caption)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption, RadMsgConvert.ToMessageBoxButton(button));
            return RadMsgConvert.ToRadMessageBoxResult(result);
        }




        public static RadMessageBoxResult Show(string message, RadMessageBoxButton button
            , RadMessageBoxIcon icon, string caption)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption
                                                    , RadMsgConvert.ToMessageBoxButton(button)
                                                    , RadMsgConvert.ToMessageBoxImage(icon));
            return RadMsgConvert.ToRadMessageBoxResult(result);

        }


        public static RadMessageBoxResult ShowDialog(string message, RadMessageBoxButton button
    , RadMessageBoxIcon icon, string caption)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption
                                                    , RadMsgConvert.ToMessageBoxButton(button)
                                                    , RadMsgConvert.ToMessageBoxImage(icon));
            return RadMsgConvert.ToRadMessageBoxResult(result);

        }


        public static RadMessageBoxResult Show(string message, RadMessageBoxButton button
            , RadMessageBoxIcon icon)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption
                                                    , RadMsgConvert.ToMessageBoxButton(button)
                                                    , RadMsgConvert.ToMessageBoxImage(icon));
            return RadMsgConvert.ToRadMessageBoxResult(result);

        }

        public static RadMessageBoxResult ShowDialog(string message, RadMessageBoxButton button
            , RadMessageBoxIcon icon)
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption
                                                    , RadMsgConvert.ToMessageBoxButton(button)
                                                    , RadMsgConvert.ToMessageBoxImage(icon));
            return RadMsgConvert.ToRadMessageBoxResult(result);

        }






        public static RadMessageBoxResult Show(string message, RadMessageBoxButton button, RadMessageBoxIcon icon, RadMessageWindowType type = RadMessageWindowType.Modal, string caption = "InstaWorkStation")
        {
            System.Windows.MessageBoxResult result = Xceed.Wpf.Toolkit.MessageBox.Show(message, caption
                                                    , RadMsgConvert.ToMessageBoxButton(button)
                                                    , RadMsgConvert.ToMessageBoxImage(icon));
            return RadMsgConvert.ToRadMessageBoxResult(result);

        }

    }
}
