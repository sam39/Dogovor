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

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Word._Application application;
            Word._Document document;

            object missingObj = System.Reflection.Missing.Value;
            object trueObj = true;
            object falseObj = false;

            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            object templatePathObj = "D:\\tmp\\test.doc"; ;

            // если вылетим не этом этапе, приложение останется открытым

            document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            //Word._Document docnew = application.Documents.Add();
            
            //{
            //    document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            //}
            //catch (Exception error)
            //{
            //    document.Close(ref falseObj, ref missingObj, ref missingObj);
            //    application.Quit(ref missingObj, ref missingObj, ref missingObj);
            //    document = null;
            //    application = null;
            //    throw error;
            //}
            application.Visible = true;

            Word.Range bookmarkRange = document.Bookmarks["aaa"].Range;
            bookmarkRange.Delete();

        }

        private void btnTemplatePath_Click(object sender, RoutedEventArgs e)
        {
            //// Create OpenFileDialog 
            //Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            //// Set filter for file extension and default file extension 
            //dlg.DefaultExt = ".png";
            //dlg.Filter = "Word Files (*.doc)|*.doc|MS Word Files (*.docx)|*.docx";


            //// Display OpenFileDialog by calling ShowDialog method 
            //Nullable<bool> result = dlg.ShowDialog();


            //// Get the selected file name and display in a TextBox 
            //if (result == true)
            //{
            //    // Open document 
            //    string filename = dlg.FileName;
            //    tbTemplatePath.Text = filename;
                
            //}
        }
    }
}