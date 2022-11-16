using System.Windows.Forms;

namespace ExcelMacroAdd.Forms.FillingOutPassportClass
{
    internal static  class МеssageView
    {
        internal static void MessageInformation(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        internal static void MessageWarning(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }

        internal static void MessageError(string textMessage, string textAttribute)
        {
            MessageBox.Show(textMessage,
                textAttribute,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}
