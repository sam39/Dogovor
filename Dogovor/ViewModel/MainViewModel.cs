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


        private ObservableCollection<Signatory> _signatorys;
        public ObservableCollection<Signatory> Signatorys
        {
            get
            {
                if (_signatorys == null)
                {
                    _signatorys = new ObservableCollection<Signatory>();
                    _signatorys.Add(new Signatory {FIO = "Новоселов Владислав Аркадьевич", Id = "Signatory1" });
                    _signatorys.Add(new Signatory { FIO = "Новоселов Эдуард Аркадьевич", Id = "Signatory2" });
                    _signatorys.Add(new Signatory { FIO = "Крылов Владислав Леонидович", Id = "Signatory3" });
                    _signatorys.Add(new Signatory { FIO = "Гераськин Ярослав Вячеславович", Id = "Signatory4" });
                    _signatorys.Add(new Signatory { FIO = "Никитина Людмила Анатольевна", Id = "Signatory5" });
                }
                return _signatorys;
            }
        }      
         
        private Contract _contract;
        public Contract Contract
        {
            get
            {
                if (_contract == null) _contract = new Contract { Currency = Currency.Доллар, Report = true, Date = DateTime.Now };
                return _contract;

            }
            set
            {
            }
                
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

        #region Выбор должности
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
            object templatePathObj = "D:\\tmp\\test.docx"; ;

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
                switch (Contract.Signatory.Id)
                {
                    case "Signatory1":
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
                    case "Signatory2":
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
                    case "Signatory3":
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
                    case "Signatory4":
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
                    case "Signatory5":
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
            return true;
        }
        #endregion Выбор должности

        private void DelBookmark(string bookmark)
        {
            Word.Range bookmarkRange;
            bookmarkRange = document.Bookmarks[bookmark].Range;
            bookmarkRange.Delete();
        }





        ////public override void Cleanup()
        ////{
        ////    // Clean up if needed

        ////    base.Cleanup();
        ////}
    }
}