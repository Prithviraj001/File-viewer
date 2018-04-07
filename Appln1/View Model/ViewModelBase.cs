using Appln1.Commands;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;

namespace Appln1.View_Model
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        #region Private Members

        private bool _CanExecute;
        private bool _WordType;
        private bool _PDFType;
        private bool _ImageType;
        private bool _VideoType;
        private bool _ExcelType;

        //private List<string> SupportedExtensions;
        private Dictionary<string, List<string>> _SupportedExtensions;

        private System.Data.DataTable _ExcelDataTable;

        private string _FileInfo;
        private string _PDFInfo;
        private string _VideoInfo;
        private string _SelectedExtension;
        private BitmapImage _ImageInfo;
        private IDocumentPaginatorSource _docText;
        private IDocumentPaginatorSource _excelText;
        string filetype;

        private ICommand _OpenDocCommand;
        private ICommand _ViewDocCommand;

        #endregion

        #region Public Members

        public BitmapImage ImageInfo
        {
            get
            {
                return _ImageInfo;
            }
            set
            {
                _ImageInfo = value;
                OnPropertyChanged("ImageInfo");
            }
        }

        public IDocumentPaginatorSource DocText
        {
            get
            {

                return _docText;
            }

            set
            {
                _docText = value;
                OnPropertyChanged("DocText");
            }
        }

        public IDocumentPaginatorSource ExcelText
        {
            get
            {

                return _excelText;
            }

            set
            {
                _excelText = value;
                OnPropertyChanged("ExcelText");
            }
        }

        public string FileInfo
        {
            get
            {
                return _FileInfo;
            }
            set
            {
                _FileInfo = value;
                OnPropertyChanged("FileInfo");
            }
        }

        public string PDFInfo
        {
            get
            {
                return _PDFInfo;
            }
            set
            {
                _PDFInfo = value;
                OnPropertyChanged("PDFInfo");
            }
        }

        public string VideoInfo
        {
            get
            {
                return _VideoInfo;
            }
            set
            {
                _VideoInfo = value;
                OnPropertyChanged("VideoInfo");
            }
        }

        public bool PDFType
        {
            get
            {
                return _PDFType;
            }
            set
            {
                _PDFType = value;
                OnPropertyChanged("PDFType");
            }
        }

        public bool ImageType
        {
            get
            {
                return _ImageType;
            }
            set
            {
                _ImageType = value;
                OnPropertyChanged("ImageType");
            }
        }

        public bool VideoType
        {
            get
            {
                return _VideoType;
            }
            set
            {
                _VideoType = value;
                OnPropertyChanged("VideoType");
            }
        }

        public bool WordType
        {
            get
            {
                return _WordType;
            }
            set
            {
                _WordType = value;
                OnPropertyChanged("WordType");
            }
        }

        public bool ExcelType
        {
            get
            {
                return _ExcelType;
            }
            set
            {
                _ExcelType = value;
                OnPropertyChanged("ExcelType");
            }
        }

        public System.Data.DataTable ExcelDataTable
        {
            get
            {
                return _ExcelDataTable;
            }
            set
            {
                _ExcelDataTable = value;
                OnPropertyChanged("ExcelDataTable");
            }
        }

        public ICommand OpenDocCommand
        {
            get
            {
                return _OpenDocCommand ?? (_OpenDocCommand = new CommandHandler((parameter) => OpenDocCommandAction(parameter), _CanExecute));
            }
        }

        public ICommand ViewDocCommand
        {
            get
            {
                return _ViewDocCommand ?? (_ViewDocCommand = new CommandHandler((parameter) => ViewDocCommandAction(parameter), _CanExecute));
            }
            
        }

        #endregion

        #region Action Commands

        private void OpenDocCommandAction(object parameter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if(openFileDialog.ShowDialog() == true)
            {
                //filename = openFileDialog.FileName;
                FileInfo = openFileDialog.FileName;
            }

        }

        private void ViewDocCommandAction(object parameter)
        {
            try
            {
                if(FileInfo != null)
                {
                    SetViewFalse();
                }

                _SelectedExtension = Path.GetExtension(FileInfo);

                filetype = GetFileType(_SelectedExtension);

                OpenVisitor opv = new OpenVisitor();

                switch (filetype)
                {
                    case "ImageFile":
                        Console.WriteLine("Image File");
                        ImageType = true;
                        Image i = new Image();
                        i.SetFilename(FileInfo);
                        i.accept(opv);
                        ImageInfo = i.bi;
                        VideoInfo = null;
                        break;

                    case "VideoFile":
                        Console.WriteLine("Video File");
                        VideoType = true;
                        Video v = new Video();
                        v.SetFilename(FileInfo);
                        v.accept(opv);
                        VideoInfo = v.info;
                        break;

                    case "PDFFile":
                        Console.WriteLine("PDF File");
                        PDFType = true;
                        PDF pdf = new PDF();
                        pdf.SetFilename(FileInfo);
                        pdf.accept(opv);
                        PDFInfo = pdf.info;
                        VideoInfo = null;
                        break;

                    case "WordFile":
                        Console.WriteLine("Word File");
                        WordType = true;
                        Word w = new Word();
                        w.SetFilename(FileInfo);
                        w.accept(opv);
                        DocText = w.info;
                        VideoInfo = null;
                        break;

                    case "ExcelFile":
                        Console.WriteLine("Excel File");
                        ExcelType = true;
                        Excel e = new Excel();
                        e.SetFilename(FileInfo);
                        e.accept(opv);
                        ExcelText = e.info;
                        VideoInfo = null;
                        break;
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Exception at ViewDocCommandAction : " + exp.Message);
            }

        }

        private string GetFileType(string SelectedExtension)
        {
            try
            {
                foreach (KeyValuePair<string, List<string>> pair in _SupportedExtensions)
                {
                    if (pair.Value.Contains(SelectedExtension))
                    {
                        return pair.Key;
                    }
                }
                return null;
            }
            catch(Exception exp)
            {
                Console.WriteLine("Exception at GetFileType : " + exp.Message);
                return null;
            }
        }

        private void InitializeExtensions()
        {
            string[] ImageExtensions = { ".png", ".jpeg", ".jpg", ".PNG", ".JPEG", ".JPG" };
            string[] VideoExtensions = { ".mp4", ".mpg", ".mpeg", ".m1v", ".mp2", ".mpa", ".mpe", ".avi",".wmv",".mkv",
            ".MP4", ".MPG", ".MPEG", ".M1V", ".MP2", ".MPA", ".MPE", ".AVI", ".WMV", ".MKV"};
            string[] ExcelExtensions = { ".xls", ".xlsx", ".csv", ".xltm", "xltx", ".XLXS", ".CSV", ".XLTM", "XLTX" };
            string[] WordExtensions = { ".doc", ".docx", ".DOC", ".DOCX" };
            string[] PDFExtensions = {".pdf", ".PDF"};
            List<string> Temp = new List<string>();
            Temp.AddRange(ImageExtensions);
            _SupportedExtensions.Add("ImageFile", Temp);
            Temp = new List<string>();
            Temp.AddRange(VideoExtensions);
            _SupportedExtensions.Add("VideoFile", Temp);
            Temp = new List<string>();
            Temp.AddRange(ExcelExtensions);
            _SupportedExtensions.Add("ExcelFile", Temp);
            Temp = new List<string>();
            Temp.AddRange(WordExtensions);
            _SupportedExtensions.Add("WordFile", Temp);
            Temp = new List<string>();
            Temp.AddRange(PDFExtensions);
            _SupportedExtensions.Add("PDFFile", Temp);
        }

        #endregion

        #region ViewModelBase Constructor
        public ViewModelBase()
        {
            _CanExecute = true;
            _FileInfo = "";
            SetViewFalse();

            _SupportedExtensions = new Dictionary<string, List<string>>();
            ImageInfo = new BitmapImage();
            ExcelDataTable = new System.Data.DataTable();
            InitializeExtensions();
        }

        private void SetViewFalse()
        {
            WordType = false;
            ExcelType = false;
            PDFType = false;
            ImageType = false;
            VideoType = false;
            _SelectedExtension = "";
        }

        #endregion

        #region OnPropertyChanged Event Handler Function

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
}
