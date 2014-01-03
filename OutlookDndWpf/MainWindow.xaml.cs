// --------------------------------------------------------------------------
// Licensed under MIT License.
//
// Outlook DnD Data Reader
// 
// File     : MainWindow.xaml.cs
// Author   : Tobias Viehweger <tobias.viehweger@yasoon.com / @mnkypete>
//
// -------------------------------------------------------------------------- 

using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OutlookDndWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        
        private void Rectangle_DragEnter(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
        }

        private void Rectangle_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
        }

        private void Rectangle_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("RenPrivateMessages"))
            {
                var dataStream = e.Data.GetData("RenPrivateMessages") as MemoryStream;

                if (dataStream != null)
                {
                    OleDataReader reader = new OleDataReader(dataStream);
                    var outlookObj = reader.ReadOutlookData();

                    this.countBox.Text = outlookObj.Items.Length.ToString();
                    this.subjectBox.Text = outlookObj.Items[0].Subject;
                    this.entryIdBox.Text = outlookObj.Items[0].EntryId;
                }
            }
        }

    }
}
