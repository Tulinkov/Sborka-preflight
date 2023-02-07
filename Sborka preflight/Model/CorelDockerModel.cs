using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Corel.Interop.VGCore;

namespace SborkaPreflight.Model
{
    class CorelDockerModel : Notifier
    {
        #region Fields
        private bool _newOrderCanExecute = true;
        private bool _listNewOrdersIsVisible = false;
        private bool _toPDFCanExecute = false;
        private bool _toPrepressCanExecute = false;
        private bool _toUserCanExecute = false;
        private bool _makeGuidesCanExecute = false;
        private bool _fitToPageCanExecute = false;
        private bool _aditionalMenuCanExecute = false;
        private string _waitingText = "";
        private int _indexOfSelectedOrder = 0;
        private Corel.Interop.VGCore.Application app;
        private WorkDate workDate;
        private OpenedOrder _openedOrder;
        private Hole _hole = new Hole();
        private Cutline _cutline = new Cutline();
        private struct Side { public ShapeRange shapes; public bool center; public bool? isHorizontal; }
        private Dictionary<cdrImageType, string> colorMode = new Dictionary<cdrImageType, string>();
        #endregion

        #region Properties
        public bool NewOrderCanExecute
        {
            get { return _newOrderCanExecute; }
            set
            {
                if (_newOrderCanExecute != value)
                {
                    _newOrderCanExecute = value;
                    OnPropertyChanged("NewOrderCanExecute");
                }
            }
        }
        public bool ListNewOrdersIsVisible
        {
            get { return _listNewOrdersIsVisible; }
            set
            {
                if (_listNewOrdersIsVisible != value)
                {
                    _listNewOrdersIsVisible = value;
                    OnPropertyChanged("ListNewOrdersIsVisible");
                }
            }
        }
        public bool ToPDFCanExecute
        {
            get { return _toPDFCanExecute; }
            set
            {
                if (_toPDFCanExecute != value)
                {
                    _toPDFCanExecute = value;
                    OnPropertyChanged("ToPDFCanExecute");
                }
            }
        }
        public bool ToPrepressCanExecute
        {
            get { return _toPrepressCanExecute; }
            set
            {
                if (_toPrepressCanExecute != value)
                {
                    _toPrepressCanExecute = value;
                    OnPropertyChanged("ToPrepressCanExecute");
                }
            }
        }
        public bool ToUserCanExecute
        {
            get { return _toUserCanExecute; }
            set
            {
                if (_toUserCanExecute != value)
                {
                    _toUserCanExecute = value;
                    OnPropertyChanged("ToUserCanExecute");
                }
            }
        }
        public bool MakeGuidesCanExecute
        {
            get { return _makeGuidesCanExecute; }
            set
            {
                if (_makeGuidesCanExecute != value)
                {
                    _makeGuidesCanExecute = value;
                    OnPropertyChanged("MakeGuidesCanExecute");
                }
            }
        }
        public bool FitToPageCanExecute
        {
            get { return _fitToPageCanExecute; }
            set
            {
                if (_fitToPageCanExecute != value)
                {
                    _fitToPageCanExecute = value;
                    OnPropertyChanged("FitToPageCanExecute");
                }
            }
        }
        public bool AditionalMenuCanExecute
        {
            get { return _aditionalMenuCanExecute; }
            set
            {
                if (_aditionalMenuCanExecute != value)
                {
                    _aditionalMenuCanExecute = value;
                    OnPropertyChanged("AditionalMenuCanExecute");
                }
            }
        }
        public string WaitingText
        {
            get { return _waitingText; }
            set
            {
                if (_waitingText != value)
                {
                    _waitingText = value;
                    OnPropertyChanged("WaitingText");
                }
            }
        }
        public int IndexOfSelectedOrder
        {
            get { return _indexOfSelectedOrder; }
            set
            {
                if (_indexOfSelectedOrder != value)
                {
                    _indexOfSelectedOrder = value;
                    OnPropertyChanged("IndexOfSelectedOrder");
                }
            }
        }
        public ObservableCollection<OrderFromFolder> ListNewOrder { get; set; }
        public OpenedOrder OpenedOrder
        {
            get { return _openedOrder; }
            set
            {
                if (_openedOrder != value)
                {
                    _openedOrder = value;
                    OnPropertyChanged("OpenedOrder");
                }
            }
        }
        public Hole Hole
        {
            get { return _hole; }
            set
            {
                if (_hole != value)
                {
                    _hole = value;
                    OnPropertyChanged("Hole");
                }
            }
        }
        public Cutline Cutline
        {
            get { return _cutline; }
            set
            {
                if (_cutline != value)
                {
                    _cutline = value;
                    OnPropertyChanged("Cutline");
                }
            }
        }
        #endregion

        #region Constructor
        public CorelDockerModel(Corel.Interop.VGCore.Application app)
        {
            this.app = app;
            workDate = new WorkDate();
            OpenedOrder = null;
            app.WindowActivate += new Corel.Interop.VGCore.DIVGApplicationEvents_WindowActivateEventHandler(WindowActivate);
            app.DocumentNew += new Corel.Interop.VGCore.DIVGApplicationEvents_DocumentNewEventHandler(DocumentNew);
            app.WindowDeactivate += new Corel.Interop.VGCore.DIVGApplicationEvents_WindowDeactivateEventHandler(WindowDeactivate);
            ListNewOrder = new ObservableCollection<OrderFromFolder>();
            colorMode.Add(cdrImageType.cdr16ColorsImage, "16 Color");
            colorMode.Add(cdrImageType.cdrBlackAndWhiteImage, "BW");
            colorMode.Add(cdrImageType.cdrCMYKColorImage, "CMYK");
            colorMode.Add(cdrImageType.cdrCMYKMultiChannelImage, "CMYK");
            colorMode.Add(cdrImageType.cdrDuotoneImage, "Duotone");
            colorMode.Add(cdrImageType.cdrGrayscaleImage, "Gray");
            colorMode.Add(cdrImageType.cdrLABImage, "LAB");
            colorMode.Add(cdrImageType.cdrPalettedImage, "Palleted");
            colorMode.Add(cdrImageType.cdrRGBColorImage, "RGB");
            colorMode.Add(cdrImageType.cdrRGBMultiChannelImage, "RGB");
            colorMode.Add(cdrImageType.cdrSpotMultiChannelImage, "Spot");
        }
        #endregion

        /// <summary>
        /// Displays a list of orders available for processing, or an error message
        /// </summary>
        public void CreateListOfOrder()
        {
            if (!Directory.Exists(PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate)))
            {
                MessageBox.Show(string.Format("Folder \"{0}\" is not found.",
                                               PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate)),
                                               "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (!MakeListNewOrders())
                MessageBox.Show(string.Format("Folder \"{0}\" doesn't contain appropriate files.",
                                               PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate)),
                                               "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            else
                ListNewOrdersIsVisible = true;
        }

        /// <summary>
        /// Publishes the layout to PDF
        /// </summary>
        public async void PublishToPDF()
        {
            await NewWaitingText("Publishing to PDF...");
            app.Optimization = true;
            string fullPathPDF = PrepressFile.GetPath(PrepressFile.pprsPath.pathPDFs, workDate);
            app.ActiveDocument.MasterPage.GuidesLayer.Printable = false;
            RemoveGuides();
            if (OpenedOrder != null)
                Checkout();
            try
            {
                if (!System.IO.Directory.Exists(fullPathPDF))
                {
                    MessageBox.Show($"Path {fullPathPDF} does not exist.", "Publish to PDF - error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                app.ActiveDocument.PDFSettings.Load(Properties.Settings.Default.PdfPreset);
                app.ActiveDocument.PDFSettings.PublishToPDF(fullPathPDF + PrepressFile.FileNameWithoutExtension(app.ActiveDocument.Title) + ".pdf");
                if (OpenedOrder != null)
                    if (app.ActiveDocument.Title == OpenedOrder.FilesList["face"])
                        OpenedOrder.isPublishedToPDF = true;
            }
            catch
            {
                MessageBox.Show("Failed to save PDF", "Publish to PDF - error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                MakeGuides();
                app.ActiveDocument.MasterPage.GuidesLayer.Printable = true;
                app.Optimization = false;
                app.Refresh();
                app.ActiveWindow.Activate();
                NewWaitingText(wait: false);
            }
            try { System.Diagnostics.Process.Start(fullPathPDF + PrepressFile.FileNameWithoutExtension(app.ActiveDocument.Title) + ".pdf"); } //Launches pdf reader
            catch { };
        }

        /// <summary>
        /// Saves the layout as v.17
        /// </summary>
        public async void SaveToV17()
        {
            await NewWaitingText("Saving as version 17...");
            app.Optimization = true;
            string fullPathCDR = PrepressFile.GetPath(PrepressFile.pprsPath.pathPDFs, workDate);
            RemoveGuides();
            if (OpenedOrder != null)
                Checkout();
            try
            {
                if (!System.IO.Directory.Exists(fullPathCDR))
                {
                    MessageBox.Show("Path " + fullPathCDR + " does not exist.", "Saving - error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                StructSaveAsOptions saveOption = new StructSaveAsOptions();
                saveOption.Version = cdrFileVersion.cdrVersion17;
                saveOption.EmbedICCProfile = false;
                app.ActiveDocument.SaveAs(fullPathCDR + PrepressFile.FileNameWithoutExtension(app.ActiveDocument.Title) + ".cdr", saveOption);
                if (OpenedOrder != null)
                    if (app.ActiveDocument.Title == OpenedOrder.FilesList["face"])
                        OpenedOrder.isPublishedToPDF = true;
            }
            catch
            {
                MessageBox.Show("Failed to save as version 17", "Saving - error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                MakeGuides();
                app.ActiveDocument.MasterPage.GuidesLayer.Printable = true;
                app.Optimization = false;
                app.Refresh();
                app.ActiveWindow.Activate();
                NewWaitingText(wait: false);
            }

        }

        /// <summary>
        /// Closes the project and moves the customer's files to the folder "Approved files"
        /// </summary>
        public async void SendToPrepress()
        {
            if (OpenedOrder != null)
                if (!OpenedOrder.isPublishedToPDF)
                    if (MessageBox.Show("The PDF-file wasn't created.\nAre you sure you want to send the order in prepress?",
                                        "Warning...", MessageBoxButton.OKCancel, MessageBoxImage.Information, MessageBoxResult.Cancel) == MessageBoxResult.Cancel)
                        return;
            await NewWaitingText("Closing document...");
            MoveFiles(PrepressFile.GetPath(PrepressFile.pprsPath.pathApproved, workDate));
            await NewWaitingText(wait: false);
            OpenedOrder.document.Close();
        }

        /// <summary>
        /// Closes the project and moves the customer's files to the folder "Problem files"
        /// </summary>
        public async void ReturnToUser()
        {
            await NewWaitingText("Closing document...");
            MoveFiles(PrepressFile.GetPath(PrepressFile.pprsPath.pathProblem, workDate));
            await NewWaitingText(wait: false);
            OpenedOrder.document.Close();
        }

        /// <summary>
        /// If Ctrl is pressed, removes guidelines.
        /// Otherwise, sets the guides to a standard offset from the edge of the page, or to 5mm if Alt is pressed
        /// </summary>
        public void MakeOrDeleteGuides()
        {
            app.Optimization = true;
            ModifierKeys keyboardModifiers = Keyboard.Modifiers;
            RemoveGuides();
            if (keyboardModifiers != ModifierKeys.Control)
                if (keyboardModifiers == ModifierKeys.Alt)
                    MakeGuides(5);
                else
                    MakeGuides();
            app.Optimization = false;
            app.Refresh();
            app.ActiveWindow.Activate();
        }

        /// <summary>
        /// Fits the layout on active page to the page
        /// </summary>
        public void FitToPage()
        {
            app.Optimization = true;
            ShapeRange shapesBlocked = FindAllShapes(app.ActiveDocument.ActivePage.Shapes);
            if (shapesBlocked.Count > 0)
                if (shapesBlocked.Shapes.FindShapes(Query: "@com.Locked").Count == 0)
                {
                    Shapes allShapes = app.ActiveDocument.ActivePage.Shapes;
                    ShapeRange shapesWithoutGuides = new ShapeRange();
                    foreach (Shape shape in allShapes)
                        if (shape.Type != cdrShapeType.cdrGuidelineShape)
                            shapesWithoutGuides.Add(shape);
                    shapesWithoutGuides.SetSize(app.ActivePage.SizeWidth, app.ActivePage.SizeHeight);
                    shapesWithoutGuides.SetPositionEx(cdrReferencePoint.cdrCenter, app.ActivePage.SizeWidth / 2, app.ActivePage.SizeHeight / 2);
                }
                else
                {
                    MessageBox.Show("The document contains locked objects. Unlock them.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            if (OpenedOrder != null)
                Checkout();
            app.Optimization = false;
            app.Refresh();
            app.ActiveWindow.Activate();
        }

        /// <summary>
        /// Starts the processing of a user-selected order
        /// </summary>
        /// <param name="arg">Indicates OK or Cancel was pressed</param>
        public async void ProcessOrder(object arg)
        {
            ListNewOrdersIsVisible = false; //hide the order list
            if (!bool.Parse((string)arg)) return; //if "Cancel" is pressed
            NewOrderCanExecute = false;
            var files = from file in Directory.EnumerateFiles(PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate),
                                                              string.Format("*({0})*.*", ListNewOrder[_indexOfSelectedOrder].FullNumber))
                        group Path.GetFileName(file)
                        by PrepressFile.NumberFromName(Path.GetFileName(file));

            if (CheckoutOrderFiles(files.First<IGrouping<string, string>>())) //file extensions verification
            {
                OpenedOrder = new OpenedOrder(files.First<IGrouping<string, string>>());
                Copy(PrepressFile.NumberFromName(OpenedOrder.FilesList["face"], "order"));
            }
            else
            {
                MessageBox.Show(string.Format("Order {0} contain non appropriate files", ListNewOrder[_indexOfSelectedOrder].FullNumber), "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            app.Optimization = true;
            await NewWaitingText("Creating new file...");
            OpenedOrder.isDigital = PrepressFile.IsDigital(OpenedOrder.FilesList["face"]);
            CreateDocument(OpenedOrder.FilesList["face"]);
            await NewWaitingText("Opening face...");
            bool isOpened = CreatePage(1); //processing of the first side
            if (OpenedOrder.FilesList.ContainsKey("back"))
            {
                await NewWaitingText("Opening back...");
                isOpened = isOpened && CreatePage(2); //processing of the second side
            }
            if (!isOpened) //in case of an error when opening
            {
                await NewWaitingText("Closing document...");
                OpenedOrder.document.Close();
                MessageBox.Show("Error opening files. The document will be closed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);            
            }
            else
            {
                OpenedOrder.document.Activate();
                OpenedOrder.document.Pages[1].Activate();
                MakeGuides();
                OpenedOrder.document.ActiveWindow.ActiveView.ToFitArea(OpenedOrder.document.Pages[1].LeftX-2,
                                                                       OpenedOrder.document.Pages[1].TopY+2,
                                                                       OpenedOrder.document.Pages[1].RightX+2,
                                                                       OpenedOrder.document.Pages[1].BottomY-2);
                Checkout(); //checking for compliance with the technical requirements
            }
            await NewWaitingText(wait: false);
            app.Optimization = false;
            app.Refresh();
            try { app.ActiveWindow.Activate(); }
            catch { };
        }

        /// <summary>
        /// Shows the setting window
        /// </summary>
        public void SettingWindow()
        {
            View.Setting settingWindow = new View.Setting();
            settingWindow.ShowDialog();
        }

        /// <summary>
        /// Copies a text information to the clipboard
        /// </summary>
        /// <param name="arg">Order or custumer number</param>
        public void Copy(object arg)
        {
            var isSuccessfully = false;
            for (int i=0; i < 10; i++) //10 attempts to copy
            {
                try
                {
                    System.Windows.Clipboard.Clear();
                    System.Windows.Clipboard.SetDataObject(arg.ToString());
                    isSuccessfully = true;
                }
                catch { };
                if (isSuccessfully) break;
                Thread.Sleep(250);
            }
            if (!isSuccessfully) MessageBox.Show("Failed to copy to the clipboard", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        /// <summary>
        /// Rechecks the layout for compliance with the technical requirements
        /// </summary>
        public void Update()
        {
            bool isOptimized = !app.Optimization;
            if (isOptimized) app.Optimization = true;
            Checkout();
            if (isOptimized)
            {
                app.Optimization = false;
                app.Refresh();
                try { app.ActiveWindow.Activate(); }
                catch { };
            }
        }

        /// <summary>
        /// Opens a customer's file
        /// </summary>
        /// <param name="arg">name of the file</param>
        public void SideOpen(object arg)
        {
            try
            {
                if (OpenedOrder.FilesList.ContainsKey(arg.ToString()))
                    Process.Start(PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate) + OpenedOrder.FilesList[arg.ToString()]);
            }
            catch
            {
                MessageBox.Show("Unable to open file");
            }            
        }

        /// <summary>
        /// Draws a circle
        /// </summary>
        public void DrawCircle()
        {
            double left = (Hole.Left == 0) ? (app.ActiveDocument.MasterPage.SizeWidth - Hole.Diameter) / 2 : Hole.Left;
            double bottom = (Hole.Bottom == 0) ? (app.ActiveDocument.MasterPage.SizeHeight - Hole.Diameter) / 2 : Hole.Bottom;
            Shape hole = app.ActiveLayer.CreateEllipse(left, bottom, left + Hole.Diameter, bottom + Hole.Diameter);
            hole.Outline.Color = app.CreateCMYKColor(100, 0, 100, 0);
            hole.Outline.Width = 0.1;
        }

        /// <summary>
        /// Prepares plotter cutting templates
        /// </summary>
        public void FixCutline()
        {
            ShapeRange cutlines = app.ActivePage.Shapes.FindShapes(Query: $"@outline.color.name = '{Cutline.ColorName}' and @outline.color.IsSpot = true");
            if (cutlines.Count > 2)
                MessageBox.Show("Caution! This document has more that two cutlines.");
            foreach (Shape shape in cutlines)
            {
                if (shape.Outline.DashDotLength != 0 && Cutline.RemoveDottedLine)
                {
                    shape.Delete();
                    continue;
                }
                shape.Outline.Color = app.CreateCMYKColor(0, 0, 0, 100);
                shape.Outline.Width = 0.1;
            }
        }

        /// <summary>
        /// Invoked when a window is activated
        /// </summary>
        /// <param name="Doc">document</param>
        /// <param name="Window">window</param>
        private void WindowActivate(Corel.Interop.VGCore.Document Doc, Corel.Interop.VGCore.Window Window)
        {
            app.ActiveDocument.Unit = cdrUnit.cdrMillimeter;
            ToPDFCanExecute = true;
            MakeGuidesCanExecute = true;
            FitToPageCanExecute = true;
            AditionalMenuCanExecute = true;
            Hole.Reset(Convert.ToInt32(Doc.MasterPage.SizeWidth), Convert.ToInt32(Doc.MasterPage.SizeHeight));
            Doc.PageChange += PageChange;
            if (OpenedOrder != null)
            {
                bool isActiveWorkingDoc = (Doc.Title == OpenedOrder.FilesList["face"]);
                ToPrepressCanExecute = isActiveWorkingDoc;
                ToUserCanExecute = isActiveWorkingDoc;
            }
        }

        /// <summary>
        /// Invoked when a new document is created
        /// </summary>
        /// <param name="Doc">document</param>
        /// <param name="FromTemplate">uses a template</param>
        /// <param name="Template">template name</param>
        /// <param name="IncludeGraphics">has include graphics</param>
        private void DocumentNew(Corel.Interop.VGCore.Document Doc, bool FromTemplate, string Template, bool IncludeGraphics)
        {
            WindowActivate(Doc, null);
        }

        /// <summary>
        /// Invoked when a window is deactivated
        /// </summary>
        /// <param name="Doc">document</param>
        /// <param name="Window">window</param>
        private void WindowDeactivate(Corel.Interop.VGCore.Document Doc, Corel.Interop.VGCore.Window Window)
        {
            ToPDFCanExecute = false;
            MakeGuidesCanExecute = false;
            FitToPageCanExecute = false;
            ToPrepressCanExecute = false;
            ToUserCanExecute = false;
            AditionalMenuCanExecute = false;
            Doc.PageChange -= PageChange;
        }

        /// <summary>
        /// Invoked when a page is changed
        /// </summary>
        /// <param name="Page">page</param>
        private void PageChange(Page Page)
        {
            Hole.Reset(Convert.ToInt32(Page.SizeHeight), Convert.ToInt32(Page.SizeHeight));
        }

        /// <summary>
        /// Creates a document based on the customer's layout
        /// </summary>
        /// <param name="name">document name</param>
        private void CreateDocument(string name)
        {
            StructCreateOptions createOpt = app.CreateStructCreateOptions();
            createOpt.Name = name;
            createOpt.Units = cdrUnit.cdrMillimeter;
            if (((OpenedOrder.Width == 322) && (OpenedOrder.Height == 450)) || //digital printing without cutting
                ((OpenedOrder.Width == 452) && (OpenedOrder.Height == 320)))
            {
                createOpt.PageWidth = OpenedOrder.Width - 2;
                createOpt.PageHeight = OpenedOrder.Height - 2;
            }
            else
            {
                int margin = 0; //for offset printing
                if (OpenedOrder.isDigital)
                    margin = 2; //for digital printing
                createOpt.PageWidth = OpenedOrder.Width + margin;
                createOpt.PageHeight = OpenedOrder.Height + margin;
            }
            createOpt.Resolution = 300;
            createOpt.ColorContext = app.CreateColorContext2("sRGB IEC61966-2.1,ISO Coated v2 (ECI),Dot Gain 15%", BlendingColorModel: clrColorModel.clrColorModelCMYK);
            OpenedOrder.document = app.CreateDocumentEx(createOpt);
            OpenedOrder.document.MasterPage.SetSize(createOpt.PageWidth, createOpt.PageHeight);         
            ((Corel.Interop.VGCore.DIVGDocumentEvents_Event)OpenedOrder.document).Close += new Corel.Interop.VGCore.DIVGDocumentEvents_CloseEventHandler(ClosingDocument);
        }

        /// <summary>
        /// Checks the layout for compliance with the technical requirements
        /// </summary>
        private void Checkout()
        {
            double shapesPage1_centerX, shapesPage1_centerY, shapesPage1_width, shapesPage1_height;
            double shapesPage2_centerX, shapesPage2_centerY, shapesPage2_width, shapesPage2_height;
            OpenedOrder.Width = OpenedOrder.document.Pages[1].SizeWidth;
            OpenedOrder.Height = OpenedOrder.document.Pages[1].SizeHeight;
            //page 1
            ShapeRange shapesPage1 = OpenedOrder.document.Pages[1].Shapes.All();
            ShapeRange shapesPage1WithoutGuidelines = new ShapeRange();
            for (int i = 1; i <= shapesPage1.Count; i++)
                if (shapesPage1[i].Type != cdrShapeType.cdrGuidelineShape)
                    shapesPage1WithoutGuidelines.Add(shapesPage1[i]); //ignore guidelines when checking
            if (shapesPage1WithoutGuidelines.Count > 0)
            {
                OpenedOrder.Page1Exist = true; //not empty page
                shapesPage1_centerX = Math.Round(shapesPage1WithoutGuidelines.CenterX, 2);
                shapesPage1_centerY = Math.Round(shapesPage1WithoutGuidelines.CenterY, 2);
                shapesPage1_width = Math.Round(shapesPage1WithoutGuidelines.SizeWidth, 2);
                shapesPage1_height = Math.Round(shapesPage1WithoutGuidelines.SizeHeight, 2);
                OpenedOrder.Page1IsCentered = ((shapesPage1_centerX == Math.Round(OpenedOrder.document.Pages[1].CenterX, 2)) && //layout alignment check
                                              (shapesPage1_centerY == Math.Round(OpenedOrder.document.Pages[1].CenterY, 2))) ? true : false;
                OpenedOrder.Page1SizeIsCorrect = ((shapesPage1_width == Math.Round(OpenedOrder.document.Pages[1].SizeWidth, 2)) && //layout size check
                                                 (shapesPage1_height == Math.Round(OpenedOrder.document.Pages[1].SizeHeight, 2))) ? 0 :
                                                 (SizeError(new Size(shapesPage1_width, shapesPage1_height), x => (x <= 2) && (x >= -4))) ? 1 : 2;
                if ((shapesPage1WithoutGuidelines.Count == 1) && (shapesPage1WithoutGuidelines[1].Type == cdrShapeType.cdrBitmapShape)) //for raster graphics
                {
                    OpenedOrder.Page1Resolution = Math.Min(shapesPage1WithoutGuidelines[1].Bitmap.ResolutionX, shapesPage1WithoutGuidelines[1].Bitmap.ResolutionY).ToString(); //resolution check
                    OpenedOrder.Page1ColorMode = colorMode[shapesPage1WithoutGuidelines[1].Bitmap.Mode]; //color mode check
                    OpenedOrder.Page1Texts = null; //text objects check
                    OpenedOrder.Page1Effects = null; //effects check
                }
                else //for vector graphics
                {
                    OpenedOrder.Page1Resolution = "0"; //resolution check
                    OpenedOrder.Page1ColorMode = "non"; //color mode check
                    OpenedOrder.Page1Texts = (FindAllShapes(OpenedOrder.document.Pages[1].Shapes).Shapes.FindShapes(Type: cdrShapeType.cdrTextShape).Count > 0) ? true : false; //text objects check
                    OpenedOrder.Page1Effects = (FindAllShapes(OpenedOrder.document.Pages[1].Shapes).Shapes.FindShapes(Query: "@com.Effects.Count > 0").Count > 0) ? true : false; //effects check
                }
            }
            else //empty page
                OpenedOrder.Page1Exist = false;
            //page 2
            ShapeRange shapesPage2 = new ShapeRange();
            ShapeRange shapesPage2WithoutGuidelines = new ShapeRange();
            if (OpenedOrder.document.Pages.Count < 2) //if the page 2 don't exist
            {
                OpenedOrder.Page2Exist = false;
                return;
            }
            shapesPage2 = OpenedOrder.document.Pages[2].Shapes.All();
            for (int i = 1; i <= shapesPage2.Count; i++)
                if (shapesPage2[i].Type != cdrShapeType.cdrGuidelineShape)
                    shapesPage2WithoutGuidelines.Add(shapesPage2[i]); //ignore guidelines when checking
            if (shapesPage2WithoutGuidelines.Count > 0)
            {
                OpenedOrder.Page2Exist = true; //not empty page
                shapesPage2_centerX = Math.Round(shapesPage2WithoutGuidelines.CenterX, 2);
                shapesPage2_centerY = Math.Round(shapesPage2WithoutGuidelines.CenterY, 2);
                shapesPage2_width = Math.Round(shapesPage2WithoutGuidelines.SizeWidth, 2);
                shapesPage2_height = Math.Round(shapesPage2WithoutGuidelines.SizeHeight, 2);
                OpenedOrder.Page2IsCentered = ((shapesPage2_centerX == Math.Round(OpenedOrder.document.Pages[2].CenterX, 2)) && //layout alignment check
                                              (shapesPage2_centerY == Math.Round(OpenedOrder.document.Pages[2].CenterY, 2))) ? true : false;
                OpenedOrder.Page2SizeIsCorrect = ((shapesPage2_width == Math.Round(OpenedOrder.document.Pages[2].SizeWidth, 2)) && //layout size check
                                                 (shapesPage2_height == Math.Round(OpenedOrder.document.Pages[2].SizeHeight, 2))) ? 0 :
                                                 (SizeError(new Size(shapesPage2_width, shapesPage2_height), x => (x <= 2) && (x >= -4))) ? 1 : 2;
                if ((shapesPage2WithoutGuidelines.Count == 1) && (shapesPage2WithoutGuidelines[1].Type == cdrShapeType.cdrBitmapShape)) //for raster graphics
                {
                    OpenedOrder.Page2Resolution = Math.Min(shapesPage2WithoutGuidelines[1].Bitmap.ResolutionX, shapesPage2WithoutGuidelines[1].Bitmap.ResolutionY).ToString(); //resolution check
                    OpenedOrder.Page2ColorMode = colorMode[shapesPage2WithoutGuidelines[1].Bitmap.Mode]; //color mode check
                    OpenedOrder.Page2Texts = null; //text objects check
                    OpenedOrder.Page2Effects = null; //effects check
                }
                else
                {
                    OpenedOrder.Page2Resolution = "0"; //resolution check
                    OpenedOrder.Page2ColorMode = "non"; //color mode check
                    OpenedOrder.Page2Texts = (FindAllShapes(OpenedOrder.document.Pages[2].Shapes).Shapes.FindShapes(Type: cdrShapeType.cdrTextShape).Count > 0) ? true : false; //text objects check
                    OpenedOrder.Page2Effects = (FindAllShapes(OpenedOrder.document.Pages[2].Shapes).Shapes.FindShapes(Query: "@com.Effects.Count > 0").Count > 0) ? true : false; //effects check
                }
            }
            else //empty page
                OpenedOrder.Page2Exist = false;
        }

        /// <summary>
        /// Creates a new page in the document and places the customer's layout there
        /// </summary>
        /// <param name="pageIndex">1 for face, 2 for back</param>
        /// <param name="shapeRange">struct Side</param>
        /// <returns></returns>
        private bool CreatePage(int pageIndex, Side shapeRange = new Side())
        {
            string side = (pageIndex == 1) ? "face" : "back";
            if ((pageIndex == 2) && (!OpenedOrder.FilesList.ContainsKey("back")) && (shapeRange.shapes == null))
                return true;
            Document document = OpenedOrder.document;
            if (document.Pages.Count < pageIndex)
                document.InsertPages(1, false, document.Pages.Count);
            Layer layer = document.Pages[pageIndex].Layers[2];
            layer.Name = side;
            if (shapeRange.shapes != null) //if shapeRange conteins shapes
                return MoveShapes(shapeRange.shapes, layer, shapeRange.center);
            if (Path.GetExtension(OpenedOrder.FilesList[side]) != ".cdr") //raster graphic
            {
                try
                {
                    StructImportOptions importOpt = app.CreateStructImportOptions();
                    importOpt.Mode = cdrImportMode.cdrImportFull;
                    importOpt.ColorConversionOptions.SourceColorProfileList = "sRGB IEC61966-2.1,ISO Coated v2 (ECI),Dot Gain 15%";
                    importOpt.ColorConversionOptions.TargetColorProfileList = "sRGB IEC61966-2.1,ISO Coated v2 (ECI),Dot Gain 15%";
                    ImportFilter importFlt = layer.ImportEx(PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate) + OpenedOrder.FilesList[side], cdrFilter.cdrAutoSense, importOpt);
                    importFlt.Finish();
                    if ((pageIndex == 1) && (document.Pages[1].Shapes.All().SizeWidth < document.Pages[1].Shapes.All().SizeHeight)) //if layout is vertical
                        {
                            OpenedOrder.ChangeWidthAndHeight();
                            document.MasterPage.SetSize(document.MasterPage.SizeHeight, document.MasterPage.SizeWidth);
                        }
                }
                catch { return false; };
            }
            else //file .cdr
            {
                Document documentCDR = OpenCDR(PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate) + OpenedOrder.FilesList[side]);
                if (documentCDR == null)
                    return false;
                documentCDR.Unit = cdrUnit.cdrMillimeter;
                int maxPages = ((pageIndex == 1) && (!OpenedOrder.FilesList.ContainsKey("back"))) ? 2 : 1;
                if (CheckoutCDR(documentCDR, maxPages))
                {
                    Side[] sides = FindSide(documentCDR, pageIndex, maxPages == 1);
                    if (sides[1].shapes != null)
                        CreatePage(2, sides[1]);
                    if (MoveShapes(sides[0].shapes, layer, sides[0].center))
                    {
                        if (((pageIndex == 1) && (sides[0].isHorizontal != true)) && //if layout is vertical
                            ((sides[0].isHorizontal == false) || (OpenedOrder.document.Pages[1].Shapes.All().SizeWidth < OpenedOrder.document.Pages[1].Shapes.All().SizeHeight)))
                            {
                                OpenedOrder.ChangeWidthAndHeight();
                                document.MasterPage.SetSize(document.MasterPage.SizeHeight, document.MasterPage.SizeWidth);
                            }
                        if (FindAllShapes(documentCDR).Count == 0) //if document is empty
                            documentCDR.Close();
                    }
                    else
                        return false;
                }
            }
            OpenedOrder.document.ActivePage.Shapes.All().AlignAndDistribute(cdrAlignDistributeH.cdrAlignDistributeHAlignCenter, //center
                                                                            cdrAlignDistributeV.cdrAlignDistributeVAlignCenter,
                                                                            cdrAlignShapesTo.cdrAlignShapesToCenterOfPage);
            return true;
        }

        /// <summary>
        /// Specifies the number and location of layouts in the customer's file
        /// </summary>
        /// <param name="document">document</param>
        /// <param name="page">page number</param>
        /// <param name="hasOneSide"> true for 4+0, false for 4+4</param>
        /// <returns>array of found sides</returns>
        private Side[] FindSide(Document document, int page, bool hasOneSide)
        {
            Side[] sides = new Side[2];
            RemoveGuides(document);
            if (hasOneSide) //if order is 4+0, all shapes are for one side
                sides[0].shapes = document.Pages[1].Shapes.All();
            else
                if (document.Pages.Count > 1)
                {
                    sides[0].shapes = document.Pages[1].Shapes.All();
                    sides[1].shapes = document.Pages[2].Shapes.All();
                }
                else
                {
                    System.Windows.Point rangeCenter = new System.Windows.Point(document.Pages[1].Shapes.All().CenterX,
                                                                                document.Pages[1].Shapes.All().CenterY);
                    ShapeRange[] shapes = { new ShapeRange(), new ShapeRange(), new ShapeRange(), new ShapeRange() };
                    for (int i = 1; i <= document.Pages[1].Shapes.Count; i++ )
                    {
                        Page pageOne = document.Pages[1];
                        if (pageOne.Shapes[i].RightX <= rangeCenter.X) //shapes for the 1 page, horizontal arrangement of the sides
                            shapes[0].Add(document.Pages[1].Shapes[i]);
                        if (pageOne.Shapes[i].LeftX >= rangeCenter.X)  //shapes for the 2 page, horizontal arrangement of the sides
                            shapes[1].Add(document.Pages[1].Shapes[i]);
                        if (pageOne.Shapes[i].BottomY >= rangeCenter.Y) //shapes for the 1 page, vertical arrangement of the sides
                            shapes[2].Add(document.Pages[1].Shapes[i]);
                        if (pageOne.Shapes[i].TopY <= rangeCenter.Y) //shapes for the 2 page, vertical arrangement of the sides
                            shapes[3].Add(document.Pages[1].Shapes[i]);
                    }
                    if (SizeError(new Size(shapes[0].SizeWidth, shapes[0].SizeHeight), x => (x <= 2) && (x >= -4)) &&
                        SizeError(new Size(shapes[1].SizeWidth, shapes[1].SizeHeight), x => (x <= 2) && (x >= -4)) &&
                        (shapes[0].Count + shapes[1].Count == document.Pages[1].Shapes.Count))
                    {
                        sides[0].shapes = shapes[0]; //shapes for the 1 page, horizontal arrangement of the sides
                        sides[1].shapes = shapes[1]; //shapes for the 2 page, horizontal arrangement of the sides
                    }
                    else
                    {
                        if (SizeError(new Size(shapes[2].SizeWidth, shapes[2].SizeHeight), x => (x <= 2) && (x >= -4)) &&
                            SizeError(new Size(shapes[3].SizeWidth, shapes[3].SizeHeight), x => (x <= 2) && (x >= -4)) &&
                            (shapes[2].Count + shapes[3].Count == document.Pages[1].Shapes.Count))
                        {
                            sides[0].shapes = shapes[2]; //shapes for the 1 page, vertical arrangement of the sides
                            sides[1].shapes = shapes[3]; //shapes for the 2 page, vertical arrangement of the sides
                        }
                        else
                            sides[0].shapes = document.Pages[1].Shapes.All(); //one side is found
                    }
                }
            for (int i = 0; i < 2; i++)
            {
                if (sides[i].shapes == null) continue; //side contains no objects
                sides[i].center = true;
                sides[i].isHorizontal = null;
                System.Windows.Rect pageRect = new System.Windows.Rect(document.Pages[1].LeftX, document.Pages[1].BottomY,
                                                                       document.Pages[1].SizeWidth, document.Pages[1].SizeHeight);
                System.Windows.Rect layoutRect = new System.Windows.Rect(sides[i].shapes.LeftX, sides[i].shapes.BottomY,
                                                                         sides[i].shapes.SizeWidth, sides[i].shapes.SizeHeight);
                if (SizeError(new Size(pageRect.Width, pageRect.Height), x => (x <= 2) && (x >= -4)))
                    if (layoutRect.IntersectsWith(pageRect))
                    {
                        sides[i].isHorizontal = (pageRect.Width > pageRect.Height) ? true : false;
                        sides[i].center = false;
                    }
            }
            return sides;
        }

        /// <summary>
        /// Checks if the size fits within the specified tolerances
        /// </summary>
        /// <param name="size">size</param>
        /// <param name="ErrorExtent">tolerances</param>
        /// <returns>true if fits, false otherwise</returns>
        private bool SizeError(Size size, Predicate<double> ErrorExtent)
        {
            double standardWidth = OpenedOrder.Width;
            double standardHeight = OpenedOrder.Height;
            return ErrorExtent(standardWidth - size.Width) && ErrorExtent(standardHeight - size.Height) ||
                   ErrorExtent(standardHeight - size.Width) && ErrorExtent(standardWidth - size.Height);
        }

        /// <summary>
        /// Moves shapes
        /// </summary>
        /// <param name="shapes">list of shapes to move</param>
        /// <param name="layer">layer on which to place</param>
        /// <param name="center">should it be centered</param>
        /// <returns>false on error</returns>
        private bool MoveShapes(ShapeRange shapes, Layer layer, bool center)
        {
            try
            {
                if (shapes.Count > 0)
                {
                    shapes.Cut();
                    layer.Paste();
                    layer.Shapes.All().Group();
                    if (center)
                        layer.Shapes.All().SetPositionEx(cdrReferencePoint.cdrCenter, layer.Page.SizeWidth / 2, layer.Page.SizeHeight / 2);
                }
                return true;
            }
            catch
            { return false; }
        }

        /// <summary>
        /// Opens a new .cdr
        /// </summary>
        /// <param name="file">file name</param>
        /// <returns>false on error</returns>
        private Document OpenCDR(string file)
        {
            Document document = FindDocument(file);
            if (document != null)
                return document;
            Process.Start(file);
            for (int i = 0; i < 5; i++)
            {
                Thread.Sleep(250);
                document = FindDocument(file);
                if (document != null)
                    return document;
            }
            return null;
        }

        /// <summary>
        /// Searches for a document among open
        /// </summary>
        /// <param name="file">file name</param>
        /// <returns>document</returns>
        private Document FindDocument(string file)
        {
            foreach (Document document in app.Documents)
                if (file.ToUpper() == string.Concat(document.FilePath, document.Name).ToUpper())
                    return document;
            return null;
        }

        /// <summary>
        /// Performs initial validation of the customer's file
        /// </summary>
        /// <param name="document">document</param>
        /// <param name="maxPage">maximum allowable number of pages</param>
        /// <returns>false if there are any problems</returns>
        private bool CheckoutCDR(Document document, int maxPage)
        {
            bool isLocked = false;
            bool isMorePages = false;
            bool isNoObjects = false;
            string resultText = string.Format("Document \"{0}\" contains:", document.Name);
            ShapeRange shapes = FindAllShapes(document);
            if (document.Pages.Count > maxPage) //page number
            {
                isMorePages = true;
                resultText += string.Format("\n- pages more than {0};", maxPage);
            }
            if (shapes.Count > 0)
            {
                if (shapes.Shapes.FindShapes(Query: "@com.Locked").Count > 0) //locked shapes
                {
                    isLocked = true;
                    resultText += "\n- locked objects";
                }
            }
            else //empty document
            {
                isNoObjects = true;
                if (resultText == string.Format("Document \"{0}\" contains:", document.Name))
                    resultText = string.Format("Document \"{0}\" is empty", document.Name);
                else
                    resultText += "\n\nDocument is empty";
            }
            resultText += "\n\nObjects won't be moved";
            bool result = isLocked || isMorePages || isNoObjects;
            if (result)
                MessageBox.Show(resultText);
            return !result;
        }

        /// <summary>
        /// Invoked when the document is closed
        /// </summary>
        private void ClosingDocument()
        {
            OpenedOrder = null;
            NewOrderCanExecute = true;
        }

        /// <summary>
        /// Determines a list of orders based on files in a folder
        /// </summary>
        /// <returns>false if orders are not found</returns>
        private bool MakeListNewOrders()
        {
            ListNewOrder.Clear();
            var groupListFiles = from file in Directory.EnumerateFiles(PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate))
                           where PrepressFile.NameIsCorrect(Path.GetFileName(file))
                           group Path.GetFileName(file)
                           by PrepressFile.NumberFromName(Path.GetFileName(file));
            foreach (var group in groupListFiles)
            {
                if (CheckoutOrderFiles(group))
                    ListNewOrder.Add(new OrderFromFolder(PrepressFile.NumberFromName(group.First<string>()),
                                                         "", PrepressFile.SizeFromName(group.First<string>())));
            }
            return (ListNewOrder.Count == 0) ? false : true;
        }

        /// <summary>
        /// Checks that the order contains all the necessary files
        /// </summary>
        /// <param name="files">list of files</param>
        /// <returns>false if the order cannot be formed from the files in stock</returns>
        private bool CheckoutOrderFiles(IGrouping<string, string> files)
        {
            bool IncorrectExtention = false;
            bool FaceExist = false;
            bool BackExist = false;
            bool TwoFace = false;
            bool TwoBack = false;
            bool indeterminateFileExist = false;
            foreach (var file in files)
            {
                string extention = Path.GetExtension(file);
                bool isFaceOrBack = false;
                if (Properties.Settings.Default.trueExtentions.IndexOf(extention, StringComparison.OrdinalIgnoreCase) < 0) //extention verification
                {
                    if (Properties.Settings.Default.admissibleExtentions.Contains(extention))
                        continue;
                    else
                        IncorrectExtention = true;
                }
                if (Path.GetFileNameWithoutExtension(file).EndsWith("-face")) //presence and number of files containing the layout of the front side
                {
                    if (!FaceExist) FaceExist = true;
                    else TwoFace = true;
                    isFaceOrBack = true;
                }
                if (Path.GetFileNameWithoutExtension(file).EndsWith("-back")) //presence and number of files containing the layout of the back side
                {
                    if (!BackExist) BackExist = true;
                    else TwoBack = true;
                    isFaceOrBack = true;
                }
                if (!isFaceOrBack) indeterminateFileExist = true; //unidentified files
            }
            return !IncorrectExtention && FaceExist && !(TwoFace || TwoBack) && !indeterminateFileExist;
        }
        
        /// <summary>
        /// Add guidelines
        /// </summary>
        /// <param name="margin">margin</param>
        private void MakeGuides(int margin = 3)
        {
            if ((OpenedOrder?.isDigital ?? false) && (margin == 3))
                margin = 5;
            ShapeRange activeShapes = app.ActiveSelectionRange; //save selection
            MakeGuide(margin, 0, 90);
            MakeGuide(app.ActivePage.SizeWidth - margin, 0, 90);
            MakeGuide(0, margin, 0);
            MakeGuide(0, app.ActivePage.SizeHeight - margin, 0);
            app.ActiveDocument.ClearSelection();
            activeShapes.AddToSelection(); //restore selection
        }

        /// <summary>
        /// Creates a guide by using one point and an angle
        /// </summary>
        /// <param name="x">Specifies the x-coordinate for the point that defines the guideline. This value is measured in document units.</param>
        /// <param name="y">Specifies the y-coordinate for the point that defines the guideline. This value is measured in document units.</param>
        /// <param name="angle">Specifies the degree to which the guideline is slanted. Values range from 0 to 360 degrees.</param>
        private void MakeGuide(double x, double y, double angle)
        {
            Shape guideLine;
            guideLine = app.ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(x, y, angle);
            guideLine.Outline.SetProperties(Color: app.CreateRGBColor(0, 0, 255));
        }

        /// <summary>
        /// Remove all guidelines
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private bool RemoveGuides(Document document = null)
        {
            document = document ?? app.ActiveDocument;
            ShapeRange GuidesShape = FindAllShapes(document).Shapes.FindShapes(Type: cdrShapeType.cdrGuidelineShape);
            foreach (Shape guide in GuidesShape)
                if (guide.Locked)
                    guide.Locked = false;
            if (GuidesShape.Count > 0)
            {
                GuidesShape.Delete();
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Finds all shapes in a document
        /// </summary>
        /// <param name="document">document name</param>
        /// <returns>list of shapes</returns>
        private ShapeRange FindAllShapes(Document document)
        {
            ShapeRange shapeRange = new ShapeRange();
            foreach (Page documentPage in document.Pages)
                shapeRange.AddRange(FindAllShapes(documentPage.Shapes));
            return shapeRange;
        }

        /// <summary>
        /// Finds all shapes in a list of shapes that are not powerclip
        /// </summary>
        /// <param name="document">document name</param>
        /// <returns>list of shapes</returns>
        private ShapeRange FindAllShapes(Shapes shapes)
        {
            ShapeRange shapeRange = shapes.FindShapes();
            foreach (Shape shapePowerClip in shapeRange.Shapes.FindShapes(Query: "!@com.powerclip.IsNull"))
                shapeRange.AddRange(FindAllShapes(shapePowerClip.PowerClip.Shapes));
            return shapeRange;
        }

        /// <summary>
        /// Changes text message
        /// </summary>
        /// <param name="text">test</param>
        /// <param name="wait">delay before UI undating</param>
        private async Task NewWaitingText(string text = "", bool wait = true)
        {
            WaitingText = text;
            if (wait)
                await Task.Run(() => { System.Threading.Thread.Sleep(300); }); //update UI
        }

        /// <summary>
        /// Moves the customer's files to the appropriate folder
        /// </summary>
        /// <param name="path">folder path</param>
        private void MoveFiles(string path)
        {
            foreach (var file in OpenedOrder.FilesList)
            {
                try
                {
                    File.Move(PrepressFile.GetPath(PrepressFile.pprsPath.pathNew, workDate) + file.Value,
                              path + file.Value);
                }
                catch { MessageBox.Show(string.Format("Failed to move file {0}", file.Value), "Error", MessageBoxButton.OK, MessageBoxImage.Warning); };
            }
        }
    }
}