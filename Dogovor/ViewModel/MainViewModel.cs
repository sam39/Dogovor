using GalaSoft.MvvmLight;
using Dogovor.Model;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Input;
using GalaSoft.MvvmLight.Command;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System;

namespace Dogovor.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// See http://www.mvvmlight.net
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        private readonly IDataService _dataService;

        //public MainViewModel()
        //{
        //    Contract 
        //}


        /// <summary>
        /// The <see cref="WelcomeTitle" /> property's name.
        /// </summary>
        public const string WelcomeTitlePropertyName = "WelcomeTitle";

        private string _welcomeTitle = string.Empty;

        /// <summary>
        /// Gets the WelcomeTitle property.
        /// Changes to that property's value raise the PropertyChanged event. 
        /// </summary>
        public string WelcomeTitle
        {
            get
            {
                return _welcomeTitle;
            }
            set
            {
                Set(ref _welcomeTitle, value);
            }
        }


        Word._Application application;
        Word._Document document;
  
         
        private Contract _contract;
        public Contract Contract
        {
            get
            {
                if (_contract == null)
                    _contract = Contract.read();
                _contract.Date = DateTime.Now;
                return _contract;
            }
            //set
            //{
            //    _contract = value;
            //}
                
        }

        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel(IDataService dataService)
        {
            _dataService = dataService;
            _dataService.GetData(
                (item, error) =>
                {
                    if (error != null)
                    {
                        // Report error here
                        return;
                    }

                    WelcomeTitle = item.Title;
                });
        }

        #region Команда Start
        RelayCommand _start;
        public ICommand Start
        {
            get
            {
                if (_start == null)
                    _start = new RelayCommand(ExecuteStartCommand, CanExecuteStartCommand);
                return _start;
            }
        }

        public void ExecuteStartCommand()
        {


            object missingObj = System.Reflection.Missing.Value;
            object trueObj = true;
            object falseObj = false;

            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            object templatePathObj = Contract.TemplatePath;

            // если вылетим не этом этапе, приложение останется открытым

            //document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            //Word._Document docnew = application.Documents.Add();
            try
            {
                document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                //document.Close(ref falseObj, ref missingObj, ref missingObj);
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                throw error;
            }
            application.Visible = true;

            Word.Range bookmarkRange;
            bookmarkRange = document.Bookmarks["Num"].Range;
            bookmarkRange.Text = Contract.Num;
            bookmarkRange = document.Bookmarks["Num1"].Range;
            bookmarkRange.Text = Contract.Num;
            bookmarkRange = document.Bookmarks["Date"].Range;
            bookmarkRange.Text = Contract.Date.ToString("dd.MM.yyyy");
            bookmarkRange = document.Bookmarks["Date1"].Range;
            bookmarkRange.Text = Contract.Date.ToString("dd.MM.yyyy");

            if (Contract.CustomerStatus == Status.Резидент)
            {
                bookmarkRange = document.Bookmarks["NotRezident"].Range;
                bookmarkRange.Delete();
            }

            if (Contract.Signatory != null)
            {
                switch (Contract.Signatory)
                {
                    case Signatory.Новоселов_В_А:
                        DelBookmark("Signatory2");
                        DelBookmark("Dover2");
                        DelBookmark("Sign2");
                        DelBookmark("Signatory3");
                        DelBookmark("Dover3");
                        DelBookmark("Sign3");
                        DelBookmark("Signatory4");
                        DelBookmark("Dover4");
                        DelBookmark("Sign4");
                        DelBookmark("Signatory5");
                        DelBookmark("Dover5");
                        DelBookmark("Sign5");
                        break;
                    case Signatory.Новоселов_Э_А:
                        DelBookmark("Signatory1");
                        DelBookmark("Dover1");
                        DelBookmark("Sign1");
                        DelBookmark("Signatory3");
                        DelBookmark("Dover3");
                        DelBookmark("Sign3");
                        DelBookmark("Signatory4");
                        DelBookmark("Dover4");
                        DelBookmark("Sign4");
                        DelBookmark("Signatory5");
                        DelBookmark("Dover5");
                        DelBookmark("Sign5");
                        break;
                    case Signatory.Крылов_В_Л:
                        DelBookmark("Signatory1");
                        DelBookmark("Dover1");
                        DelBookmark("Sign1");
                        DelBookmark("Signatory2");
                        DelBookmark("Dover2");
                        DelBookmark("Sign2");
                        DelBookmark("Signatory4");
                        DelBookmark("Dover4");
                        DelBookmark("Sign4");
                        DelBookmark("Signatory5");
                        DelBookmark("Dover5");
                        DelBookmark("Sign5");
                        break;
                    case Signatory.Гераськин_Я_В:
                        DelBookmark("Signatory1");
                        DelBookmark("Dover1");
                        DelBookmark("Sign1");
                        DelBookmark("Signatory2");
                        DelBookmark("Dover2");
                        DelBookmark("Sign2");
                        DelBookmark("Signatory3");
                        DelBookmark("Dover3");
                        DelBookmark("Sign3");
                        DelBookmark("Signatory5");
                        DelBookmark("Dover5");
                        DelBookmark("Sign5");
                        break;
                    case Signatory.Никитина_Л_А:
                        DelBookmark("Signatory1");
                        DelBookmark("Dover1");
                        DelBookmark("Sign1");
                        DelBookmark("Signatory2");
                        DelBookmark("Dover2");
                        DelBookmark("Sign2");
                        DelBookmark("Signatory3");
                        DelBookmark("Dover3");
                        DelBookmark("Sign3");
                        DelBookmark("Signatory4");
                        DelBookmark("Dover4");
                        DelBookmark("Sign4");
                        break;
                }
            }
        }

        public bool CanExecuteStartCommand()
        {
            if (Contract.TemplatePath != String.Empty)
                return true;
            else return false;
        }
        #endregion Команда Start

        private void DelBookmark(string bookmark)
        {
            Word.Range bookmarkRange;
            bookmarkRange = document.Bookmarks[bookmark].Range;
            bookmarkRange.Delete();
        }


        #region Команда Save
        RelayCommand _save;
        public ICommand Save
        {
            get
            {
                if (_save == null)
                    _save = new RelayCommand(ExecuteSaveCommand, CanExecuteSaveCommand);
                return _save;
            }
        }

        public void ExecuteSaveCommand()
        {
            Contract.save();
        }

        public bool CanExecuteSaveCommand()
        {
            return true;
        }
        #endregion Команда Save


        #region Команда SetTemplate
        RelayCommand _setTemplate;
        public ICommand SetTemplate
        {
            get
            {
                if (_setTemplate == null)
                    _setTemplate = new RelayCommand(ExecuteSetTemplateCommand, CanExecuteSetTemplateCommand);
                return _setTemplate;
            }
        }

        public void ExecuteSetTemplateCommand()
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".docx";
            dlg.Filter = "Word Files (*.doc)|*.doc|MS Word Files (*.docx)|*.docx";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                Contract.TemplatePath = filename;
                RaisePropertyChanged(() => Contract);
            }
        }

        public bool CanExecuteSetTemplateCommand()
        {
            return true;
        }
        #endregion Команда SetTemplate


        ////public override void Cleanup()
        ////{
        ////    // Clean up if needed

        ////    base.Cleanup();
        ////}
    }
}