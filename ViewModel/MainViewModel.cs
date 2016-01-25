//  Example code only, feel free to copy and re-use.

using System;
using Ookii.Dialogs.Wpf;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using System.Windows.Input;
using System.Windows;
using System.IO;
using System.Linq;
using System.IO.Compression;
using Microsoft.VisualBasic.FileIO;
using WordUPdater.HelperClasses;
using System.Windows.Forms;
using WordUPdater.Views;
using System.Windows.Controls;
using System.Windows.Interactivity;
using System.Collections.Generic;

namespace WordUpdater.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// See http://www.galasoft.ch/mvvm
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        #region Construction
        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        /// 
        public MainViewModel()
        {
            UpdatePictures = new RelayCommand(() => UpdatePicturesExecute(), () => true);
            ExitApplication = new RelayCommand(() => ExitApplicationExecute(), () => true);
            SelectImagesFolder = new RelayCommand(() => SelectImagesFolderExecute(), () => true);
            SelectWordFile = new RelayCommand(() => SelectWordFileExecute(), () => true);
            SaveBitmapFromClipboard = new RelayCommand(() => SaveBitmapFromClipboardExecute(), () => true);
            ShowHelp = new RelayCommand(() => ShowHelpExecute(), () => true);
            RenameFiles = new RelayCommand(() => RenameFilesExecute(), () => true);
            ZipPath = null;
            Image = null;
            FileToReplace = "";

        }


        #endregion

        #region Fields

        string _systemMessage;
        string _imagesPath;
        int _numberOfImagesInWord;
        string _docxPath;
        string _backupFolder;
        string _picFromClipboardPath;

        string ZipName;
        string ZipPath;
        string Image;
        const string imageString = "\\Image";

        string _fileToReplace;


        #endregion

        #region Members

        public string FileToReplace
        {
            get
            {
                return _fileToReplace;
            }
            set
            {
                if (_fileToReplace == value)
                    return;
                _fileToReplace = value;
                RaisePropertyChanged("FileToReplace");
            }
        }

        public string SystemMessage
        {
            get
            {
                return _systemMessage;
            }
            set
            {
                if (_systemMessage == value)
                    return;
                _systemMessage = value;
                RaisePropertyChanged("SystemMessage");
            }
        }

        public int NumberOfImagesInWord
        {
            get
            {
                return _numberOfImagesInWord;
            }
            set
            {
                if (_numberOfImagesInWord == value)
                    return;
                _numberOfImagesInWord = value;
                RaisePropertyChanged("NumberOfImagesInWord");
            }
        }

        public string ImagesPath
        {
            get { return _imagesPath; }
            set
            {
                if (_imagesPath == value)
                    return;
                _imagesPath = value;
                RaisePropertyChanged("ImagesPath");
            }
        }

        public string DocxPath
        {
            get { return _docxPath; }
            set
            {
                if (_docxPath == value)
                    return;
                _docxPath = value;
                RaisePropertyChanged("DocxPath");
            }
        }

        public string BackupFolder
        {
            get { return _backupFolder; }
            set
            {
                if (_backupFolder == value)
                    return;
                _backupFolder = value;
                RaisePropertyChanged("BackupFolder");
            }
        }

        public string DocxName { get; set; }

        public string PicFromClipboardPath
        {
            get
            { return _picFromClipboardPath; }
            set
            {
                if (_picFromClipboardPath == value)
                    return;
                _picFromClipboardPath = value;
                RaisePropertyChanged("PicFromClipboardPath");
            }
        }

        #endregion

        #region Icommands

        public ICommand UpdatePictures { get; private set; } //prevent from running when there is no file selected

        public ICommand ExitApplication { get; private set; }

        public ICommand SelectImagesFolder { get; private set; }

        public ICommand SelectWordFile { get; private set; }

        public ICommand SaveBitmapFromClipboard { get; set; }

        public ICommand ShowHelp { get; set; }

        public ICommand RenameFiles { get; set; }


        #endregion

        #region Methods

        #region ExecuteMethods

        private void RenameFilesExecute()
        {
            //Check whether the paths are selected.
            if (CheckImagepath() == false)
            { return; }

            RenameFilesMethod(ImagesPath);
        }

        private void SelectImagesFolderExecute()
        {
            OpenFolder();
        }
        private void SelectWordFileExecute()
        {
            OpenWordFile();
        }
        private void ExitApplicationExecute()
        {
            System.Windows.Application.Current.Shutdown();
        }
        private void SaveBitmapFromClipboardExecute()
        {
            SaveBTMFromClipboard();
        }
        private void UpdatePicturesExecute()
        {
            //reset the property to 0
            NumberOfImagesInWord = 0;

            //Check whether the paths are selected.
            if (Checkpaths() == false)
            { return; }

            //Assign the Name of the Zip file and rename docx To Zip
            RenameDocxToZip();

            //create the backupFolderPath
            MakeBackupFolderPath();

            // call subroutine to read all the png in the zip file and create a back up.
            imageReader(ZipPath, BackupFolder);

            //create entry in zip file equal to the number of word file pictures.
            CreateEntryInZipFIle();

            //Rename the Zip file back to Docx.
            RenameZipToDocx();

        }
        private void ShowHelpExecute()
        {
            HelpWindow wnd = new HelpWindow();
            wnd.Show();
            //System.Windows.MessageBox.Show("Application Made by Spyros, please call for help or bug declaration. Every feedback is appreciated!");
        }
        #endregion

        #region Open Files and Folder methods

        private void RenameFilesMethod(string dirpath)
        {
            //string dirpath = @"C:\Users\spyros.magripis\Pictures\LOS_rainbow";
            string backupDir = dirpath + @"\backup";
            Directory.CreateDirectory(backupDir);
            int exitValue = 9; // make room before this picture
            DirectoryInfo d = new DirectoryInfo(dirpath);
            FileInfo[] infos = d.GetFiles();
            string[] backupList = Directory.GetFiles(dirpath); // make a back up list
            int fCount = Directory.GetFiles(dirpath, "*", System.IO.SearchOption.TopDirectoryOnly).Length; //get number of files FIX IT TO GET THE LARGEST NAME OF THE FILE
            int flag = 0;
            // copy the pic files    
            foreach (string f in backupList)
            {
                // Remove Path from the file name.
                string fName = f.Substring(dirpath.Length + 1);
                File.Copy(Path.Combine(dirpath, fName), Path.Combine(backupDir, fName), true);
            }
            foreach (FileInfo f in infos.Reverse())
            {
                File.Move(f.FullName, f.FullName.ToString().Replace((fCount - flag).ToString(), (fCount - flag + 1).ToString()));
                flag += 1;
                if (fCount - flag == exitValue - 2)
                {
                    break;
                }
            }
        }

        private void OpenFolder()
        {
            var BrowseFolder = new VistaFolderBrowserDialog();
            BrowseFolder.RootFolder = Environment.SpecialFolder.MyDocuments;
            BrowseFolder.ShowDialog();
            ImagesPath = BrowseFolder.SelectedPath;
        }
        public void OpenWordFile()
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".docx"; // Default file extension
            dlg.Filter = "Text documents (.docx)|*.docx"; // Filter files by extension

            // Show open file dialog box
            bool? result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                DocxPath = dlg.FileName;
                DocxName = Path.GetFileName(dlg.FileName);

            }
            else
            {
                DisplayMessage("There was no docx file specified, please try again");
                return;
            }
        }
        #endregion

        #region ZipMethods
        public string replaceDocxWithZip(string extension)
        {
            return extension.Replace(".docx", ".zip");

        }
        private void imageReader(string zipPath_, string extractImagesPath)
        {

            //Read files in a folder
            using (ZipArchive archive = ZipFile.OpenRead(zipPath_))
            {
                if (!System.IO.Directory.Exists(extractImagesPath))
                {
                    System.IO.Directory.CreateDirectory(extractImagesPath);
                }
                else
                {
                    foreach (string foundFile in Microsoft.VisualBasic.FileIO.FileSystem.GetFiles(extractImagesPath, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.*"))
                    {
                        Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(foundFile, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently);
                    }
                }

                // make a backup copy of the zip in the backup folder and rename it to back to zip
                File.Copy(zipPath_, extractImagesPath + "\\" + Path.GetFileName(zipPath_).Replace(".zip", ".docx"));

                // extract entry to the backup folder
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase) || entry.FullName.EndsWith(".gif", StringComparison.OrdinalIgnoreCase) || entry.FullName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || entry.FullName.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase) || entry.FullName.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
                    {
                        entry.ExtractToFile(Path.Combine(extractImagesPath, entry.Name));
                        NumberOfImagesInWord += 1;
                    }
                }
                DisplayMessage("number of pictures in docx is " + NumberOfImagesInWord.ToString());
            }
        }
        private void CreateEntryInZipFIle()
        {
            //sort alphabetically
            int i = 1;
            List<string> list = new List<string>();
            foreach (string file in Directory.EnumerateFiles(ImagesPath))
            {
                list.Add(file);
            }
            var ImagesList = new List<string>();
            ImagesList.AddRange(list);

            ImagesList.Sort();

            foreach (var file in ImagesList)
            {

                if (Path.GetExtension(file).EndsWith(".png") || Path.GetExtension(file).EndsWith(".gif") || Path.GetExtension(file).EndsWith(".jpg") || Path.GetExtension(file).EndsWith(".bmp") || Path.GetExtension(file).EndsWith(".jpeg"))
                {

                    Image = (Path.GetFullPath(file));

                    imageIntoZip(ZipPath, Image, i); //create an entry in the zip file    
                    i += 1;

                }
            }
        }
        #endregion

        #region BMPFromClipboardMethod

        private void SaveBTMFromClipboard()
        {

            // Check whether imagesPath and clipboard are defined/empty.
            if (CheckForImagePathAndClipboard() == false)
            { return; }

            //check if the option to modify specific file is selected 
            switch (FileToReplace)
            {
                case "":
                    int i = 0;
                    int a = 0;

                    //count pics in the imagespath
                    exportToBMPLib Counter = new exportToBMPLib();

                    foreach (string file in Directory.EnumerateFiles(ImagesPath))
                    {
                        if (Path.GetExtension(file).EndsWith(".png") || Path.GetExtension(file).EndsWith(".gif") || Path.GetExtension(file).EndsWith(".jpg") || Path.GetExtension(file).EndsWith(".bmp") || Path.GetExtension(file).EndsWith(".jpeg"))
                        {
                            // use a Python like slice function to get an approx latest 10 chars to filter the Image numbering
                            string Alphanumstr = PythonSlice.Slice(file, file.Length - 10, file.Length, 1); // consider changing 10 based on the length of the image file extension
                            // get the numbers from an alphanumeric string
                            string NumericString = System.Text.RegularExpressions.Regex.Match(Alphanumstr, @"\d+").Value;

                            bool result = Int32.TryParse(NumericString, out i);
                            if (result)
                            {
                                if (i > a) // check if the last reported value is the largest.
                                {
                                    a = i;
                                }
                            }
                            else
                            {
                                DisplayMessage("Check the name of a file, it should end with a numeric value");
                                return;
                            }
                        }
                    }

                    PicFromClipboardPath = ImagesPath + imageString + (a + 1).ToString("D3") + ".bmp";
                    Counter.exportToBMP(PicFromClipboardPath);
                    DisplayMessage("Image " + (a + 1).ToString("D3") + " created");
                    break;


                default:
                    exportToBMPLib oImgObj = new exportToBMPLib();
                    PicFromClipboardPath = ImagesPath + imageString + int.Parse(FileToReplace).ToString("D3") + ".bmp";
                    oImgObj.exportToBMP(PicFromClipboardPath);
                    DisplayMessage("Image " + int.Parse(FileToReplace).ToString("D3") + " created");
                    break;
            }


        }
        #endregion

        #region Helper Functions/Methods
        /// <summary>
        /// Creates a Backup Folder breaking down the name folder into year-month-day-hours.minutes.seconds .
        /// </summary>
        /// <returns></returns>
        private string MakeBackupFolderPath()
        {
            return BackupFolder = SpecialDirectories.MyDocuments + "\\WordUpdaterBackupFolder\\" + DateTime.Now.ToString("yyyy-MM-dd-hh.mm.ss");
        }
        private bool Checkpaths()
        {
            if (string.IsNullOrEmpty(DocxPath))
            {
                DisplayMessage("Select Word File");
                return false;
            }
            if (string.IsNullOrEmpty(ImagesPath))
            {
                DisplayMessage("Select the Folder containing the Images");
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool CheckImagepath()
        {
            if (string.IsNullOrEmpty(ImagesPath))
            {
                DisplayMessage("Select the Folder containing the Images");
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool CheckForImagePathAndClipboard()
        {
            if (string.IsNullOrEmpty(ImagesPath))
            {
                DisplayMessage("Select the Folder containing the Images");
                return false;
            }

            if (System.Windows.Forms.Clipboard.GetDataObject() != null)
            {
                System.Windows.Forms.IDataObject oDataObj = System.Windows.Forms.Clipboard.GetDataObject();
                if (!(oDataObj.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap)))
                {

                    DisplayMessage("Clipboard is empty, please try again");
                    return false;
                }
                else
                    return true;
            }
            else
                return true;
        }
        private void RenameDocxToZip()
        {
            //assign the zip file name equal to the docx's
            ZipName = replaceDocxWithZip(DocxName);
            ZipPath = Path.Combine(Path.GetDirectoryName(DocxPath), ZipName);

            //rename to zip file
            FileSystem.RenameFile(DocxPath, ZipName);
        }
        /// <summary>
        /// Copies and image file into the selected zipFile.
        /// </summary>
        /// <param name="zipPath_"></param>
        /// <param name="image"></param>
        /// <param name="imageNR"></param>
        public void imageIntoZip(string zipPath_, string image, int imageNR)
        {

            using (ZipArchive archive = ZipFile.Open(zipPath_, ZipArchiveMode.Update))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    //renames file to be copied in zip to the corresponding file into the zip
                    if (entry.Name.Equals("image" + imageNR.ToString() + ".png"))
                    //if (entry.FullName.EndsWith(imageNR.ToString() + ".png"))
                    {
                        entry.Delete();
                        archive.CreateEntryFromFile(image, "word\\media\\image" + imageNR.ToString() + ".png");
                        break; // exits the loop otherwise the collection number is changing and thats an error
                    }
                    else if (entry.Name.EndsWith(imageNR.ToString() + ".gif"))
                    {
                        entry.Delete();
                        archive.CreateEntryFromFile(image, "word\\media\\image" + imageNR.ToString() + ".gif");
                        break; // exits the loop otherwise the collection number is changing and thats an error
                    }
                    else if (entry.Name.EndsWith(imageNR.ToString() + ".jpeg"))
                    {
                        entry.Delete();
                        archive.CreateEntryFromFile(image, "word\\media\\image" + imageNR.ToString() + ".jpeg");
                        break; // exits the loop otherwise the collection number is changing and thats an error
                    }
                    else if (entry.Name.EndsWith(imageNR.ToString() + ".jpg"))
                    {
                        entry.Delete();
                        archive.CreateEntryFromFile(image, "word\\media\\image" + imageNR.ToString() + ".jpg");
                        break; // exits the loop otherwise the collection number is changing and thats an error
                    }
                    else if (entry.Name.EndsWith(imageNR.ToString() + ".epf"))
                    {
                        entry.Delete();
                        archive.CreateEntryFromFile(image, "word\\media\\image" + imageNR.ToString() + ".epf");
                        break; // exits the loop otherwise the collection number is changing and thats an error
                    }
                }
            }
        }
        /// <summary>
        /// Displays the system message at the bottom of the main window.
        /// </summary>
        /// <param name="message"></param>
        public void DisplayMessage(string message)
        {
            SystemMessage += Microsoft.VisualBasic.DateAndTime.Now.ToString() + ":     " + message + Environment.NewLine;
        }

        /// <summary>
        /// Renames a zip file to docx file.
        /// </summary>
        private void RenameZipToDocx()
        {
            try
            {
                FileSystem.RenameFile(ZipPath, DocxName); // rename zip to docx
            }
            //{
            //    System.IO.Directory.Delete(PAthWithoutFileName(DocxPath) + tempString, true);
            //}
            catch (Exception ex)
            {
                DisplayMessage(ex.Message.ToString());
            }
        }
        #endregion

        #region Command Can Execute Functions

        #endregion







        #endregion






    }
}

namespace MyAttachedBehaviors
{
    /// <summary>
    ///     Intent: Behavior which means a scrollviewer will always scroll down to the bottom.
    /// </summary>
    public class AutoScrollBehavior : Behavior<ScrollViewer>
    {
        private double _height = 0.0d;
        private ScrollViewer _scrollViewer = null;

        protected override void OnAttached()
        {
            base.OnAttached();

            this._scrollViewer = base.AssociatedObject;
            this._scrollViewer.LayoutUpdated += new EventHandler(_scrollViewer_LayoutUpdated);
        }

        private void _scrollViewer_LayoutUpdated(object sender, EventArgs e)
        {
            if (Math.Abs(this._scrollViewer.ExtentHeight - _height) > 1)
            {
                this._scrollViewer.ScrollToVerticalOffset(this._scrollViewer.ExtentHeight);
                this._height = this._scrollViewer.ExtentHeight;
            }
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();

            if (this._scrollViewer != null)
            {
                this._scrollViewer.LayoutUpdated -= new EventHandler(_scrollViewer_LayoutUpdated);
            }
        }
    }
}
