using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Corel.Interop.VGCore;


namespace SborkaPreflight.Model
{
    #region class WordDate
    /// <summary>
    /// Helper class for working with dates
    /// </summary>
    class WorkDate
    {
        private DateTime date;
        public string Day { get { return date.ToString("dd"); } }
        public string Month { get { return date.ToString("MM"); } }
        public string Year { get { return date.ToString("yyyy"); } }

        public WorkDate()
        {
            date = DateTime.Now;
        }
    }
    #endregion

    #region class PrepressFile
    static class PrepressFile
    {
        public enum pprsPath { pathNew, pathPDFs, pathApproved, pathProblem };

        /// <summary>
        /// Returns the path to a working folder
        /// </summary>
        /// <param name="pathType">folder type</param>
        /// <param name="date">date</param>
        /// <returns>path</returns>
        static public string GetPath(pprsPath pathType, WorkDate date)
        {
            string path = Properties.Settings.Default[pathType.ToString()].ToString();
            return BindVariableData(path, date);
        }

        /// <summary>
        /// Replaces variables in a string with a date
        /// </summary>
        /// <param name="str">source string</param>
        /// <param name="date">date</param>
        /// <returns>new string with inserted date</returns>
        static private string BindVariableData(string str, WorkDate date)
        {
            Dictionary<string, string> variableData = new Dictionary<string, string>();
            variableData.Add("[day]", "{0}");
            variableData.Add("[month]", "{1}");
            variableData.Add("[year]", "{2}");
            foreach (var keyPair in variableData)
                str = str.Replace(keyPair.Key, keyPair.Value);
            str = string.Format(str, date.Day, date.Month, date.Year);
            return str;
        }

        /// <summary>
        /// Removes an extension from a string
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>file name without extension</returns>
        static public string FileNameWithoutExtension(string fileName)
        {
            fileName = fileName.Trim().Trim('.');
            int pointPosition = fileName.IndexOf('.');
            if (pointPosition != -1)
                fileName = fileName.Remove(pointPosition);
            return fileName;
        }

        /// <summary>
        /// Checks the file name
        /// </summary>
        /// <param name="str">file name</param>
        /// <returns>true if correct, false otherwise</returns>
        static public bool NameIsCorrect(string str)
        {
            string pattern = Properties.Settings.Default.regexFullNumberInFileName;
            return Regex.IsMatch(str, pattern);
        }

        /// <summary>
        /// Gets the order number, customer number, or both from the file name
        /// </summary>
        /// <param name="str">file name</param>
        /// <param name="type">number type</param>
        /// <returns>order number, customer number, or both</returns>
        static public string NumberFromName(string str, string type = "full")
        {
            string pattern = Properties.Settings.Default.regexFullNumberInFileName;
            return (Regex.Matches(str, pattern)[0].Groups[type].Success) ? Regex.Matches(str, pattern)[0].Groups[type].Value
                                                                         : Regex.Matches(str, pattern)[0].Groups[1].Value;
        }

        /// <summary>
        /// Method stub.
        /// Determines what type of printing will be used for the order
        /// </summary>
        /// <param name="str">file name</param>
        /// <returns>always true</returns>
        static public bool IsDigital(string str)
        {
            //old code, used before purchasing digital presses
            //bool containsOffset = str.Contains("_offset");
            //string pattern = Properties.Settings.Default.regexPrintrunInFileName;
            //Match matchPrintRun = Regex.Match(str, pattern);
            //bool printRun = matchPrintRun.Success ? ((int.Parse(matchPrintRun.Groups["printrun"].ToString()) <= 500) ? true : false ) : false;
            //return !containsOffset || printRun;
            return true;
        }

        /// <summary>
        /// Gets the order size from the file name
        /// </summary>
        /// <param name="str">file name</param>
        /// <returns>size or null if not found</returns>
        static public Size? SizeFromName(string str)
        {
            string pattern = Properties.Settings.Default.regexSizeInFileName;
            if (!Regex.IsMatch(str, pattern))
                return null;
            string sizeFromName = Regex.Matches(str, pattern)[0].Groups[1].Value;
            string[] size = sizeFromName.Split(new char[] {'x', 'X', 'х', 'Х'});
            return (int.Parse(size[0]) > int.Parse(size[1])) ? new Size(int.Parse(size[0]), int.Parse(size[1])) : new Size(int.Parse(size[1]), int.Parse(size[0]));
        }
    }
    #endregion

    #region class OrderFromFolder
    class OrderFromFolder
    {
        public string Name { get; set; }
        public string FullNumber { get; set; }
        public string Size {
            get
            {
                return (_size == null) ? "undefined" : string.Format("{0}x{1} mm", _size.Value.Width, _size.Value.Height);
            }
        }

        private Size? _size;

        public OrderFromFolder(string fullNumber, string name, Size? size)
        {
            FullNumber = fullNumber;
            Name = name;
            _size = size;
        }
    }
    #endregion

    #region class OpenedOrder
    class OpenedOrder : Notifier
    {
        #region Fields
        public Document document = null;
        private double _width = 90;
        private double _height = 50;
        private int _numberCustomer;
        private int _numberOrder;
        private bool _page1Exist;
        private bool _page1IsCentered;
        private int _page1SizeIsCorrect;
        private int _page1Resolution;
        private string _page1ColorMode;
        private bool? _page1Texts;
        private bool? _page1Effects;
        private bool _page2Exist;
        private bool _page2IsCentered;
        private int _page2SizeIsCorrect;
        private int _page2Resolution;
        private string _page2ColorMode;
        private bool? _page2Texts;
        private bool? _page2Effects;
        public bool isPublishedToPDF = false;
        public bool isDigital;
        //private Hole _hole;
        #endregion

        #region Properties
        public Dictionary<string, string> FilesList { get; set; }
        public double Width
        {
            get { return _width; }
            set
            {
                if (_width != value)
                {
                    _width = value;
                    OnPropertyChanged("Width");
                    OnPropertyChanged("DocumentSize");
                }
            }
        }
        public double Height
        {
            get { return _height; }
            set
            {
                if (_height != value)
                {
                    _height = value;
                    OnPropertyChanged("Height");
                    OnPropertyChanged("DocumentSize");
                }
            }
        }
        public int NumberCustomer
        {
            get
            { return _numberCustomer; }
            set
            {
                if (_numberCustomer != value)
                {
                    _numberCustomer = value;
                    OnPropertyChanged("NumberCustomer");
                }
            }
        }
        public int NumberOrder
        {
            get { return _numberOrder; }
            set
            {
                if (_numberOrder != value)
                {
                    _numberOrder = value;
                    OnPropertyChanged("NumberOrder");
                }
            }
        }
        public string Colored
        {
            get { return string.Format("{0}+{1}", (_page1Exist) ? 4 : 0, (_page2Exist) ? 4 : 0); }
        }
        public string DocumentSize
        {
            get { return string.Format("{0}x{1} mm", _width, _height); }
        }
        public bool Page1Exist
        {
            get { return _page1Exist; }
            set
            {
                if (_page1Exist != value)
                {
                    _page1Exist = value;
                    OnPropertyChanged("Page1Exist");
                    OnPropertyChanged("Colored");
                }
            }
        }
        public string Page1Extension
        {
            get { return string.Format(" ({0})", Path.GetExtension(FilesList["face"])); }
        }
        public bool Page1IsCentered
        {
            get { return _page1IsCentered; }
            set
            {
                if (_page1IsCentered != value)
                {
                    _page1IsCentered = value;
                    OnPropertyChanged("Page1IsCentered");
                }
            }
        }
        public int Page1SizeIsCorrect //0 - correct, 1 - norm, 2 - error;
        {
            get { return _page1SizeIsCorrect; }
            set
            {
                if (_page1SizeIsCorrect != value)
                {
                    _page1SizeIsCorrect = value;
                    OnPropertyChanged("Page1SizeIsCorrect");
                }
            }
        }
        public int Page1ResolutionIsCorrect //0 - correct, 1 - norm, 2 - error;
        {
            get { if (_page1Resolution > 279) return 0;
                  if (_page1Resolution > 249) return 1;
                  if (_page1Resolution == 0) return 3;
                  return 2;
            }
        }
        public string Page1Resolution
        {
            get { return (_page1Resolution == 0) ? "" : string.Format("{0} dpi", _page1Resolution); }
            set
            {
                if (_page1Resolution != int.Parse(value))
                {
                    _page1Resolution = int.Parse(value);
                    OnPropertyChanged("Page1Resolution");
                    OnPropertyChanged("Page1ResolutionIsCorrect");
                }
            }
        }
        public string Page1ColorMode
        {
            get { return _page1ColorMode; }
            set
            {
                if (_page1ColorMode != value)
                {
                    _page1ColorMode = value;
                    OnPropertyChanged("Page1ColorMode");
                }
            }
        }
        public bool? Page1Texts
        {
            get { return _page1Texts; }
            set
            {
                if (_page1Texts != value)
                {
                    _page1Texts = value;
                    OnPropertyChanged("Page1Texts");
                }
            }
        }
        public bool? Page1Effects
        {
            get { return _page1Effects; }
            set
            {
                if (_page1Effects != value)
                {
                    _page1Effects = value;
                    OnPropertyChanged("Page1Effects");
                }
            }
        }
        public bool Page2Exist
        {
            get { return _page2Exist; }
            set
            {
                if (_page2Exist != value)
                {
                    _page2Exist = value;
                    OnPropertyChanged("Page2Exist");
                    OnPropertyChanged("Colored");
                }
            }
        }
        public string Page2Extension
        {
            get { return (FilesList.ContainsKey("back")) ? string.Format(" ({0})", Path.GetExtension(FilesList["back"])) : ""; }
        }
        //public string Page1Coordinate
        //{
        //    get { return _page1Coordinate; }
        //    set
        //    {
        //        if (_page1Coordinate != value)
        //        {
        //            _page1Coordinate = value;
        //            OnPropertyChanged("Page1Coordinate");
        //        }
        //    }
        //}
        public bool Page2IsCentered
        {
            get { return _page2IsCentered; }
            set
            {
                if (_page2IsCentered != value)
                {
                    _page2IsCentered = value;
                    OnPropertyChanged("Page2IsCentered");
                }
            }
        }
        public int Page2SizeIsCorrect //0 - correct, 1 - norm, 2 - error;
        {
            get { return _page2SizeIsCorrect; }
            set
            {
                if (_page2SizeIsCorrect != value)
                {
                    _page2SizeIsCorrect = value;
                    OnPropertyChanged("Page2SizeIsCorrect");
                }
            }
        }
        public int Page2ResolutionIsCorrect //0 - correct, 1 - norm, 2 - error;
        {
            get
            {
                if (_page2Resolution > 279) return 0;
                if (_page2Resolution > 249) return 1;
                if (_page2Resolution == 0) return 3;
                return 2;
            }
        }
        public string Page2Resolution
        {
            get { return (_page2Resolution == 0) ? "" : string.Format("{0} dpi", _page2Resolution); }
            set
            {
                if (_page2Resolution != int.Parse(value))
                {
                    _page2Resolution = int.Parse(value);
                    OnPropertyChanged("Page2Resolution");
                    OnPropertyChanged("Page2ResolutionIsCorrect");
                }
            }
        }
        public string Page2ColorMode
        {
            get { return _page2ColorMode; }
            set
            {
                if (_page2ColorMode != value)
                {
                    _page2ColorMode = value;
                    OnPropertyChanged("Page2ColorMode");
                }
            }
        }
        public bool? Page2Texts
        {
            get { return _page2Texts; }
            set
            {
                if (_page2Texts != value)
                {
                    _page2Texts = value;
                    OnPropertyChanged("Page2Texts");
                }
            }
        }
        public bool? Page2Effects
        {
            get { return _page2Effects; }
            set
            {
                if (_page2Effects != value)
                {
                    _page2Effects = value;
                    OnPropertyChanged("Page2Effects");
                }
            }
        }
        //public Hole Hole
        //{
        //    get { return _hole; }
        //    set
        //    {
        //        if (_hole != value)
        //        {
        //            _hole = value;
        //            OnPropertyChanged("Hole");
        //        }
        //    }
        //}
        #endregion

        #region Constructor
        public OpenedOrder(IGrouping<string, string> files)
        {
            FilesList = new Dictionary<string,string>();
            int fileNonUsing = 1;
            foreach (var file in files)
            {
                if (Properties.Settings.Default.admissibleExtentions.Contains(Path.GetExtension(file)))
                {
                    FilesList.Add(string.Format("nonUsed{0}", fileNonUsing), file);
                    fileNonUsing++;
                    continue;
                }
                if (Path.GetFileNameWithoutExtension(file).EndsWith("-face"))
                    FilesList.Add("face", file);
                if (Path.GetFileNameWithoutExtension(file).EndsWith("-back"))
                    FilesList.Add("back", file);
            }
            Size? size = PrepressFile.SizeFromName(FilesList["face"]);
            if (size != null)
            {
                Width = (int)size.Value.Width + 2;
                Height = (int)size.Value.Height + 2;
            }
            NumberCustomer = int.Parse(PrepressFile.NumberFromName(FilesList["face"], "customer"));
            NumberOrder = int.Parse(PrepressFile.NumberFromName(FilesList["face"], "order"));
            //Hole = new Hole((int)size.Value.Width, (int)size.Value.Height);
        }
        #endregion

        public void ChangeWidthAndHeight()
        {
            double temp = Width;
            Width = Height;
            Height = temp;
        }
    }
    #endregion

    #region class Hole
    class Hole : Notifier
    {
        private int _diameter = 5;
        private double _left;
        private double _right;
        private double _top;
        private double _bottom;
        private int _productWidth;
        private int _productHeight;

        public int Diameter
        {
            get { return _diameter; }
            set
            {
                if (_diameter != value)
                {
                    _diameter = value;
                    Right = (Left == 0) ? 0 : _productWidth - _diameter - Left;
                    Bottom = (Top == 0) ? 0 : _productHeight - _diameter - Top;
                    OnPropertyChanged("Diameter");
                }
            }
        }
        public double Left
        {
            get { return _left; }
            set
            {
                if (_left != value)
                {
                    _left = value;
                    Right = (value == 0) ? 0 : _productWidth - _diameter - value;
                    OnPropertyChanged("Left");
                }
            }
        }
        public double Right
        {
            get { return _right; }
            set
            {
                if (_right != value)
                {
                    _right = value;
                    Left = (value == 0) ? 0 : _productWidth - _diameter - value;
                    OnPropertyChanged("Right");
                }
            }
        }
        public double Top
        {
            get { return _top; }
            set
            {
                if (_top != value)
                {
                    _top = value;
                    Bottom = (value == 0) ? 0 : _productHeight - _diameter - value;
                    OnPropertyChanged("Top");
                }
            }
        }
        public double Bottom
        {
            get { return _bottom; }
            set
            {
                if (_bottom != value)
                {
                    _bottom = value;
                    Top = (value == 0) ? 0 : _productHeight - _diameter - value;
                    OnPropertyChanged("Bottom");
                }
            }
        }

        public void Reset(int width, int height)
        {
            _productWidth = width;
            _productHeight = height;
            Diameter = 5;
            _left = -19999;
            Left = 0;
            _top = -19999;
            Top = 7;
        }
    }
    #endregion   

    #region class Clipboard
    public static class Clipboard
    {
        [System.Runtime.InteropServices.DllImport("user32", SetLastError = true)]
        static extern bool OpenClipboard(System.IntPtr WinHandle);
        [System.Runtime.InteropServices.DllImport("user32", SetLastError = true)]
        static extern bool EmptyClipboard();
        [System.Runtime.InteropServices.DllImport("user32", SetLastError = true)]
        static extern bool CloseClipboard();
        [System.Runtime.InteropServices.DllImport("user32", SetLastError = true)]
        static extern bool SetClipboardData(uint uFormat, IntPtr data);
        public static void CopyTextToClipboard(string text)
        {
            if (OpenClipboard(System.IntPtr.Zero))
            {
                EmptyClipboard();
                SetClipboardData(13, Marshal.StringToHGlobalUni(text));
                CloseClipboard();
            }
        }
    }
    #endregion

    #region class Cutline
    class Cutline : Notifier
    {
        private string _colorName = "Cut";
        private bool _removeDottedLine = true;

        public string ColorName
        {
            get { return _colorName; }
            set
            {
                if (value != _colorName)
                {
                    _colorName = value;
                    OnPropertyChanged("ColorName");
                }
            }
        }
        public bool RemoveDottedLine
        {
            get { return _removeDottedLine; }
            set
            {
                if (value != _removeDottedLine)
                {
                    _removeDottedLine = value;
                    OnPropertyChanged("RemoveDottedLine");
                }
            }
        }
    }
    #endregion
}
