using System.Windows;
using Dogovor.ViewModel;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System;

namespace Dogovor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            Closing += (s, e) => ViewModelLocator.Cleanup();
        }

    }
}