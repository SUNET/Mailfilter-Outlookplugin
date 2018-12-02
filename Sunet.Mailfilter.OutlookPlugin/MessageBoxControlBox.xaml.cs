using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MessageBox = System.Windows.Forms.MessageBox;

namespace Sunet.Mailfilter.OutlookPlugin
{
    //MessageBoxControlBox addon created by Marcus Liljebergh
    //Email: Marcus@Liljebergh.se
    //
    //
    //Contact for further details!
    public partial class MessageBoxControlBox : Window
    {

        public string Reason { get; set; }

        public MessageBoxControlBox()
        {
            InitializeComponent();
            btcSend.IsEnabled = false;
        }

        private void BtcSend_Click(object sender, RoutedEventArgs e)
        {
            if (System.Windows.Forms.MessageBox.Show("Send selected mail with comment?", "Send", MessageBoxButtons.YesNo) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }
            this.Reason = tbMail.Text;
            this.DialogResult = true;
            this.Close();
        }

        private void BtcCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }


        private void TbMail_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (tbMail.Text != String.Empty)
            {
                btcSend.IsEnabled = true;
            }
            else
            {
                btcSend.IsEnabled = false;
            }
        }
    }
}
