using System;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;


namespace test
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    /*
    
    AUTHOR: Emma Sevant (emma.sevant@jacobs.com)

    KEY:
     !!        :  Pending Code Changes
     // OLD    :  Old line of code, should be deleted once new code functions 

<<<<<<< HEAD
    SUMMARY:
    This script gives the application it's functionality, it is written in C#. 
    
    The user can input properties in each textbox/ drop down/ date selector. These properties can be 
    saved so that the user can fill in the properties in stages. The properties are stored in a .txt 
    file which is automatically generated and named 'BXXXXXXX-XXXX-Options (DD/MM/YY/).txt' where 
    BXXXXXXX is the project code, XXXX is the diversion number, and the date is the date when the 
    file is generated. This file is updated when the user saves. The report will be generated in the 
    same folder as this file.
=======
    User Enters Properties.
    -> boxColor() changes the boxes on each row from red to green when they are populated

    Save/ saveas buttons.
    -> this writes each property into a txt file which is named and saved in a chosen folder by the user

    Open button.
    -> User selcts a txt file 
    -> Parameters are read from txt file then atomatiaclly populated
    -> This also changes the file path in directory.ToolTip

    Create Report button.
    -> 

    User closes.
    -> !! should propmt the user to save before closing if they havn't saved 


>>>>>>> parent of 40739c0 (Practice Commit)

    When the user is ready, they can click on the create report button. This copys the report 
    template to the same folder as the properties file and names it 'BXXXXXXX-XXXX-1001_A (DD/MM/YY/)'. 
    The application will save the name of this report to the .txt file so that it can be updated. The 
    selected properties are then populated throught the report using word's custom properties function.
    
    */

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent(); 
            save.BorderBrush = new SolidColorBrush(Colors.Orange);
        }


        // Important Buttons and associated Functions
        private void packTxt(string associatedReportFilePath, string propertiesTxtfileName)
        {
            // Write all user inputs into a txt file

            // Checks how many properties there are in the app (numberOfProperties)
            int numberOfProperties = 1;
            while (FindName($"name{numberOfProperties}") != null | numberOfProperties == 16)
            { numberOfProperties = numberOfProperties + 1; }
            numberOfProperties = numberOfProperties - 1;

            // Column for property name(defined by label content)
            string[] propertyName = new string[numberOfProperties + 3];
            // Column for property value(defined by textbox/ combobox/ datepicker values)
            string[] propertyValue = new string[numberOfProperties + 3]; 
            
            propertyName[0] = "// List of the report's Custom Properties and their entered values";
            if (associatedReportFilePath == null) 
            { propertyName[1] = "// - No Report Linked";}
            else
            { propertyName[1] = "// - " + associatedReportFilePath; }

            int i = 1;
            while (i <= numberOfProperties)
            {
                if (i != 16)
                {
                    Label lb = FindName($"name{i}") as Label;
                    propertyName[i+2] = (string)lb.Content;
                    object isTextBox = FindName($"text{i}");
                    object isComboBox = FindName($"select{i}");
                    object isDatePicker = FindName($"date{i}");
                    if (isTextBox != null)
                    {
                        TextBox tb = FindName($"text{i}") as TextBox;
                        if (tb.Text != "")
                            propertyValue[i+2] = tb.Text;
                        else { propertyValue[i+2] = "-empty-"; }
                    }
                    else if (isComboBox != null)
                    {
                        ComboBox cb = FindName($"select{i}") as ComboBox;
                        if (cb.Text != "")
                            propertyValue[i + 2] = cb.Text;
                        else { propertyValue[i + 2] = "-empty-"; }
                    }
                    else if (isDatePicker != null)
                    {
                        DatePicker dp = FindName($"date{i}") as DatePicker;
                        if (dp.SelectedDate.ToString() != "")
                            propertyValue[i + 2] = dp.SelectedDate.ToString();
                        else { propertyValue[i + 2] = "-empty-"; }
                    } 
                }
                
                i++;
            }

            //// Write first Column
            File.WriteAllLines(propertiesTxtfileName, propertyName);

            // Write second Column and add it onto the fist in the txt file
            var file = File.ReadAllLines(propertiesTxtfileName);
            for (int ii = 0; ii < file.Length; ii++)
                file[ii] += '\t' + propertyValue[ii]; //add the second column after the first, with a tab
            File.WriteAllLines(propertiesTxtfileName, file);
        }
        private string[] sortPropertiesTxtFile(string propertiesFilePath, out string[] outputPropertyArray)
        {
            // read input values file
            string txtFileString = System.IO.File.ReadAllText(propertiesFilePath);

            if (txtFileString.StartsWith("// List of the report's Custom Properties and their entered values"))
            {
                // split up the string into an array (-> txtFileMessyArray) and delete un-necessary rows (-> txtFileArray)
                string[] txtFileArray = txtFileString.Split('\t', '\n', '\r');
                string[] propertiesArray = new string[txtFileArray.Length];
                
                // Exclude empty cells
                var ii = 0; // propertiesArray index
                for (int i = 9; i < txtFileArray.Length; i++)
                    if (txtFileArray[i].Length != 0)
                    { propertiesArray[ii] = txtFileArray[i]; ii++; }


                string line3 = txtFileArray[3];
                directory.ToolTip = propertiesFilePath + "\n" + line3.Substring(5);
                if (line3.Contains("No Report Linked") == false)
                {
                    createButton.Content = "Update Report!";
                    createButton.ToolTip = "Update CDS Report";
                }
                outputPropertyArray = propertiesArray;
                return outputPropertyArray;
            }
            else 
            {
                MessageBox.Show("Wrong txt file type\nFile should be in the format: BXXXXXXX-XXXX-Options (DD/MM/YY).txt");
                directory.ToolTip = "A folder is not currently selected; save as, open, or create report to select one";
                string[] emptyPropertiesArray = new string[] { null };
                return outputPropertyArray = emptyPropertiesArray;
            }
            
        }
        private void unpackTxt(string propertiesFilePath)
        {
            string[] propertiesArray;
            sortPropertiesTxtFile(propertiesFilePath, out propertiesArray);

            if (propertiesArray[0] != null) // Make sure correct file type was opened
            {
                for (int i = 0; i < propertiesArray.Length; i = i + 2)
                {
                    if ((i + 2) / 2 == 16) { i = i + 2; }
                    int n = (i + 2) / 2;
                    if (FindName($"select{n}") != null)
                    {
                        ComboBox tb = FindName($"select{n}") as ComboBox;
                        tb.Text = propertiesArray[i + 1];
                    }
                    else if (FindName($"text{n}") != null)
                    {
                        TextBox tb = FindName($"text{n}") as TextBox;
                        if (propertiesArray[i + 1] != "-empty-")
                        { tb.Text = propertiesArray[i + 1]; }
                        else
                        { tb.Text = ""; }
                    }
                    else if (FindName($"date{n}") != null)
                    {
                        DatePicker dp = FindName($"date{n}") as DatePicker;
                        if (propertiesArray[i + 1] != "-empty-")
                        { dp.SelectedDate = Convert.ToDateTime(propertiesArray[i + 1]); } 
                        else
                        {
                            dp.SelectedDate = null;
                            boxColor(dp);
                        }
                    }
                }
            }
        }
        private bool saveasProperties()
        {
            createButton.Content = "Create Report!";
            createButton.ToolTip = "Create CDS Report";

            string date = DateTime.Now.ToString();
            date = date.Replace('/', '.');
            date = date.Remove(date.Length - 9, 9);
            string projNo;
            string diversion;
            if (text9.Text != "") { projNo = text9.Text; }
            else { projNo = "BXXXXXXX"; }
            if (text7.Text != "") { diversion = text7.Text; }
            else { diversion = "XXXX"; }
            string initialFileName = projNo + "-" + diversion + "-Options (" + date + ")";

            // Open save as dialog and describe the dialog options (i,e. save as a txt file, start in AutomatedDocument folder, etc)
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = @"\\GBBHM1-FIL003\Projects\";
            saveFileDialog.FileName = initialFileName;
            saveFileDialog.Filter = "Text Files | *.txt";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == true)
            {
                directory.ToolTip = saveFileDialog.FileName + "\nNo Report Linked";

                // Create txt file using inputValuesTxt function (defined above)
                packTxt("No Report Linked", saveFileDialog.FileName);
                save.BorderBrush = new SolidColorBrush(Colors.Orange);
                MessageBox.Show("Options Saved-as!");
                return true;
            }
            else
            { MessageBox.Show("Could not save");  return false; }
                

        }
        private void saveProperties()
        {
            try
            {
                string date = DateTime.Now.ToString();
                date = date.Replace('/', '.');
                date = date.Remove(date.Length - 9, 9);
                string projNo;
                string diversion;
                if (text9.Text != "") { projNo = text9.Text; }
                else { projNo = "BXXXXXXX"; } // !! add a check for length of project no
                if (text7.Text != "") { diversion = text7.Text; }
                else { diversion = "XXXX"; }
                string updatedFileName = projNo + "-" + diversion + "-Options (" + date + ").txt";

                string directoryString = (string)directory.ToolTip;
                string divider = "\n";
                string propertiesFilePath;
                string reportFilePath;
                propertiesFilePath = directoryString.Substring(0, directoryString.IndexOf(divider));
                reportFilePath = directoryString.Substring(directoryString.IndexOf(divider) + divider.Length);
                
                if (directoryString.Contains("A properties file is not currently selected"))
                { saveasProperties(); }
                else 
                {
                    string propertiesFileName = propertiesFilePath.Substring(propertiesFilePath.Length - 38);
                    string propertiesFolder = propertiesFilePath.Substring(0, propertiesFilePath.Length - 38);
                    if (updatedFileName != propertiesFileName)
                    {
                        System.IO.File.Move(propertiesFolder + propertiesFileName, propertiesFolder + updatedFileName);
                        directory.ToolTip = propertiesFolder + updatedFileName + "\n" + reportFilePath;
                        propertiesFilePath = propertiesFolder + updatedFileName;
                    }

                    // Check if file already exists. If yes, overwirte it
                    if (File.Exists(propertiesFilePath))
                    {
                        packTxt(reportFilePath, propertiesFilePath);
                        save.BorderBrush = new SolidColorBrush(Colors.Orange);

                        MessageBox.Show("Options Saved!");
                    }
                    else { MessageBox.Show("Properites file path doesn't exist!\nCould'nt save file"); }
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
            }

        }
        private void save_Click(object sender, RoutedEventArgs e)
        { saveProperties(); }
        private void saveas_Click(object sender, RoutedEventArgs e)
        { saveasProperties(); }
        private void Open_Click(object sender, RoutedEventArgs e)
        {
            // need to prompt user to open file then save file location to directory icon ToolTip
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = @"\\GBBHM1-FIL003\Projects\";
            openFileDialog.Filter = "Text Files | *.txt";
            try
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    unpackTxt(openFileDialog.FileName);
                }
            }
            catch { MessageBox.Show("Couldnt open!" + "\n" + "Please connect to global protect"); }
            
        }
        private void Window_Closed(object sender, EventArgs e) 
        {
            if (((SolidColorBrush)save.BorderBrush).Color == Colors.Red)
            {
                if (MessageBox.Show("Options Not Saved. \nDo you want to save before closing?",
                                      "Confirmation", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                { saveProperties(); }
            }
        }


        // Update report
        private void createButton_Click(object sender, RoutedEventArgs e)
        {
            saveProperties();
            String templateReportFilePath = @"\\GBBHM1-FIL003\Admin\Gas\03 Software\Report Automation App\CDS_Blank.docx";

            // To copy a file to another location and
            // overwrite the destination file if it already exists.

            string directoryString = (string)directory.ToolTip;
            string txtFilePath = directoryString.Substring(0, directoryString.IndexOf("\n"));

            string reportFilePath = directoryString.Substring(directoryString.IndexOf("\n") + 1);
            if (reportFilePath.Contains("No Report Linked"))
            {
                string newReportFolder = txtFilePath.Remove(txtFilePath.Length - 38, 38);
                string date = DateTime.Now.ToString();
                date = date.Replace('/', '.');
                date = date.Remove(date.Length - 9, 9);
                string projNo;
                string diversion;
                if (text9.Text != "") { projNo = text9.Text; }
                else { projNo = "BXXXXXXX"; }
                if (text7.Text != "") { diversion = text7.Text; } 
                else { diversion = "XXXX"; }
                string reportFileName = projNo + "-" + diversion + "-1001_A (" + date + ")";
                reportFilePath = newReportFolder + "\\" + reportFileName + ".docx";

                if (File.Exists(reportFilePath) == false)
                {
                    directory.ToolTip = txtFilePath + "\n" + reportFilePath;
                    saveProperties();
                    System.IO.File.Copy(templateReportFilePath, reportFilePath, true);
                }
                else { MessageBox.Show(reportFilePath + " already exists!"); } 
                MessageBox.Show("This may take a minute, message box will pop up when done"); // !! change to loading message
            }
            else
            { 
                if (File.Exists(reportFilePath) == false)
                {
                    MessageBox.Show("Report Name Changed!\nPlease select report");
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.InitialDirectory = @"\\GBBHM1-FIL003\Projects\";
                    openFileDialog.Filter = "Word Documents|*.docx";
                    try
                    {
                        if (openFileDialog.ShowDialog() == true)
                        {
                            reportFilePath = openFileDialog.FileName;
                            directory.ToolTip = txtFilePath + "\n" + openFileDialog.FileName;
                            saveProperties();
                        }
                        else { MessageBox.Show("No file seleted!"); return; }
                    }
                    catch { MessageBox.Show("Couldnt open!" + "\n" + "Please connect to global protect"); }
                }
                MessageBox.Show("This may take a minute, message box will pop up when done"); // !! change to loading message
            }
            // Update word document
            try { updateDoc(reportFilePath); } 
            finally { ReleaseComObjectsUsingGC(); }

            // Update excel documents
            //{ updateEx(); } // !! add crossing register stage here
        }
        private void updateDoc(string reportFilePath)
        {
            try
            {
                if (File.Exists(reportFilePath))
                {
                    bool wasOpen;
                    if (fileIsOpen(reportFilePath))
                    {
                        wasOpen = true;
                        try { closeDoc(reportFilePath); }
                        finally { ReleaseComObjectsUsingGC(); }
                    }
                    else { wasOpen = false; }


                    // Update name of report and properties (incase a different date or project no/ diversion number entered)
                    string directoryString = (string)directory.ToolTip;
                    string txtFilePath = directoryString.Substring(0, directoryString.IndexOf("\n"));

                    string newReportFolder = txtFilePath.Remove(txtFilePath.Length - 38, 38);
                    string date = DateTime.Now.ToString();
                    date = date.Replace('/', '.');
                    date = date.Remove(date.Length - 9, 9);
                    string projNo;
                    string diversion;
                    if (text9.Text != "") { projNo = text9.Text; }
                    else { projNo = "BXXXXXXX"; }
                    if (text7.Text != "") { diversion = text7.Text; } 
                    else { diversion = "XXXX"; }
                    string updatedFileName = projNo + "-" + diversion + "-1001_A (" + date + ").docx";
                    string reportFileName = reportFilePath.Substring(reportFilePath.Length - 38);

                    if (updatedFileName != reportFileName)
                    {
                        reportFilePath = newReportFolder + updatedFileName;
                        System.IO.File.Move(newReportFolder + reportFileName, reportFilePath);
                        directory.ToolTip = txtFilePath + "\n" + reportFilePath;
                        saveProperties();
                    }

                    // Code from: https://social.msdn.microsoft.com/Forums/sqlserver/en-US/8dc4afdf-8d12-4b6e-8de8-fc990f39c8c1/creating-n-accessing-custombuiltin-document-properties?forum=worddev

                    // Define perameters for later
                    object missing = Missing.Value;
                    object DocCustomProps;

                    // Define aDoc as document and wordApp as the word application
                    Word.Application wordApp;
                    Word._Document aDoc;

                    // Open and activate word doc containing custom propeties
                    wordApp = new Word.Application();
                    aDoc = wordApp.Documents.Open(reportFilePath, ReadOnly: false);
                    aDoc = wordApp.ActiveDocument;
                    aDoc.Application.Options.WarnBeforeSavingPrintingSendingMarkup = false;

                    //Get the CustomDocumentProperties collection and find out type.
                    DocCustomProps = aDoc.CustomDocumentProperties;
                    Type typeDocCustomProps = DocCustomProps.GetType();

                    // read input values file
                    directoryString = (string)directory.ToolTip;
                    string propertiesTxtfileName = directoryString.Substring(0, directoryString.IndexOf("\n"));

                    // Form txt properties file into an array
                    string[] custPropertyArray;
                    sortPropertiesTxtFile(propertiesTxtfileName, out custPropertyArray);

                    string customProperty;
                    string customPropValue;

                    // Allocate relevent property names to word custom property in the document
                    for (int i = 0; i < 30 - 1; i = i + 2) // !! i < 30  needs to be changed when more properties are added
                    {
                        if (i != 2 && i != 6 && i != 18 && i != 24 && i != 28)  // !! make sure all label names = a customProperty in the word doc
                        {
                            customProperty = custPropertyArray[i];
                            customPropValue = custPropertyArray[i + 1];

                            try
                            {
                                typeDocCustomProps.InvokeMember("Item",
                                                   BindingFlags.Default |
                                                   BindingFlags.SetProperty,
                                                   null, DocCustomProps,
                                                   new object[] { customProperty, customPropValue });
                            }
                            catch { MessageBox.Show($"Custom property does not exist in word template for label: '{customProperty}'"); }
                        }
                    }


                    // Update Property Fields in Document
                    aDoc.Fields.Update();

                    //Save the document  
                    aDoc.Save();

                    if (wasOpen == true)
                    { wordApp.Visible = true; }
                    else
                    {
                        aDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                        wordApp.Quit();

                        if (wordApp != null)
                        {
                            if (Marshal.IsComObject(wordApp))
                            {
                                int outstanding_refs = Marshal.ReleaseComObject(wordApp);
                            }
                        }
                        if (aDoc != null)
                        {
                            if (Marshal.IsComObject(aDoc))
                            {
                                int outstanding_refs = Marshal.ReleaseComObject(aDoc);
                            }
                        }
                    }

                    MessageBox.Show("File Updated!");
                    return;
                }
                else { MessageBox.Show("Document doesnt exist!"); return; }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void closeDoc(string reportFilePath)
        {
            // Saves and Closes report: "reportFilePath"

            // Find the report from the users' open documents
            Word.Application app = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            if (app == null)
            { MessageBox.Show("Couldn't Connect to Word."); return;}

            string shortReportFilePath = reportFilePath.Substring(reportFilePath.Length - 38);
            // loop through users' open documents to find the correct one
            foreach (Word.Document d in app.Documents)
            {
                string shortFullName = d.FullName.Substring(d.FullName.Length - 38); 
                shortFullName = shortFullName.Replace('/', '\\');

                // if the document is the report, save it can close it
                if (shortFullName.ToLower() == shortReportFilePath.ToLower())
                {
                    object saveOption = Word.WdSaveOptions.wdSaveChanges; 
                    object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
                    object routeDocument = false;
                    d.Close(ref saveOption, ref originalFormat, ref routeDocument); 

                    // make sure application and document COM objects are released
                    if (app != null)
                    {
                        if (Marshal.IsComObject(app))
                        {int outstanding_refs = Marshal.ReleaseComObject(app);}
                    }
                    if (d != null)
                    {
                        if (Marshal.IsComObject(d))
                        {int outstanding_refs = Marshal.ReleaseComObject(d);}
                    }
                }
            }
        }
        private void updateEx()
        {
            // Define filePath, where the document CDS_PROPS is saved
            // !! add crossing register template to admin folder (with the report template)
            String filePath = @"C:\Users\sevantej\OneDrive - Jacobs\Documents\Technical\Report Automation\Assumptions.xlsx";

            try
            {
                // Make sure file exists
                if (File.Exists(filePath))
                {

                    // Define perameters for later
                    object missing = Missing.Value;
                    object DocCustomProps;

                    // Define aDoc as document and wordApp as the word application
                    Excel.Application oXL;
                    Excel._Workbook oWB;


                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    oXL.Visible = true;

                    //open existing Excel file
                    oWB = oXL.Workbooks.Open(filePath, FileMode.Open, FileAccess.Read);

                    //get Sheet
                    Excel.Worksheet oSheet = (Excel.Worksheet)oWB.Worksheets[2];


                    //// Get a new workbook.
                    //oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                    //oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                    //Get the CustomDocumentProperties collection and find out type.
                    DocCustomProps = oWB.CustomDocumentProperties;
                    Type typeDocCustomProps = DocCustomProps.GetType();

                    string strIndex = "Client Name";
                    string strValue = select1.Text;

                    oSheet.Cells[100, 1] = strIndex;
                    oSheet.Cells[100, 2] = strValue;

                    //Save the document  
                    //oWB.Save();

                    oXL.Quit();
                    MessageBox.Show("Excel Saved!");


                    return;
                }
                else
                {
                    MessageBox.Show("Document doesnt exist!");
                    return;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public bool fileIsOpen(string filePath)
        {
            System.IO.FileStream a = null;

            try
            {
                a = System.IO.File.Open(filePath,
                System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None);
                return false;
            }
            catch (System.IO.IOException ex)
            {
                return true;
            }

            finally
            {
                if (a != null)
                {
                    a.Close();
                    a.Dispose();
                }
            }
        }
        static public void ReleaseComObjectsUsingGC()
        {
            // COM & garbage collection help from Brian Boye: https://github.com/People-Places-Solutions/Geo-Digital_ExcelWrapper 
            
            /*
            The generally accepted best practice is not to force a garbage collection 
            in the majority of cases; however, you can release COM objects using the
            .Net garbage collector, as long as there are no references to the objects. 
            In other words, the objects are set to null.
            Be aware that GC.Collect can be a time consuming process depending 
            on the number of objects.
            You would also need to call GC.Collect and GC.WaitForPendingFinalizers twice 
            when working with Office COM objects since the first time you call the methods 
            we only release objects that we are not referencing with our own variables.
            The second time the two methods are called is because the RCW for each COM object 
            needs to run a finalizer that actually fully removes the COM Object from memory.
            So, it is totally acceptable to see the following code in you COM add-in projects:
            */
                GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return;
        }


        // Scroller/ button options
        private void buttonColor(string buttonName)
        {
            // Function used to change the 'selected' buttons colour to blue and the rest to white
            rdButton.Background = new SolidColorBrush(Colors.Black);
            phButton.Background = new SolidColorBrush(Colors.Black);
            dmButton.Background = new SolidColorBrush(Colors.Black);
            cButton.Background = new SolidColorBrush(Colors.Black);
            eButton.Background = new SolidColorBrush(Colors.Black);
            dButton.Background = new SolidColorBrush(Colors.Black);
            rdButton.Foreground = new SolidColorBrush(Colors.White);
            phButton.Foreground = new SolidColorBrush(Colors.White);
            dmButton.Foreground = new SolidColorBrush(Colors.White);
            cButton.Foreground = new SolidColorBrush(Colors.White);
            eButton.Foreground = new SolidColorBrush(Colors.White);
            dButton.Foreground = new SolidColorBrush(Colors.White);

            if (buttonName == "rd") { rdButton.Background = new SolidColorBrush(Colors.Gray); }
            else if (buttonName == "ph") { phButton.Background = new SolidColorBrush(Colors.Gray); }
            else if (buttonName == "dm") { dmButton.Background = new SolidColorBrush(Colors.Gray); }
            else if (buttonName == "c") { cButton.Background = new SolidColorBrush(Colors.Gray); }
            else if (buttonName == "e") { eButton.Background = new SolidColorBrush(Colors.Gray); }
            else if (buttonName == "d") { dButton.Background = new SolidColorBrush(Colors.Gray); }
        }
        private void boxColor(object sender)
        {
            // Change the colour of the save button border when properties are changed (will revert to orange once saved)
            save.BorderBrush = new SolidColorBrush(Colors.Red);

            // This function changes the relevent box from red to green when the user inputs data

            // Sender is the object containing all the details about the...
            // TextBox/ ComboBox/ DatePicker that has been pressed

            // Define if sender is TextBox/ ComboBox/ DatePicker, then find its' Name
            string selectName = "";
            if (sender is TextBox)
            { selectName = ((TextBox)sender).Name; }
            else if (sender is ComboBox)
            { selectName = ((ComboBox)sender).Name; }
            else if (sender is DatePicker)
            { selectName = ((DatePicker)sender).Name; }
            else
            { MessageBox.Show("sender type not identified"); }

            if (selectName != "")
            {
                // Find the variable number and therefore the relevent box
                int n = (int)Char.GetNumericValue(selectName[selectName.Length - 1]);

                if (Char.IsNumber(selectName, selectName.Length - 3) == true && Char.IsNumber(selectName, selectName.Length - 2) == true)
                {
                    // three digit numbers
                    int l = (int)Char.GetNumericValue(selectName[selectName.Length - 2]);
                    int m = (int)Char.GetNumericValue(selectName[selectName.Length - 2]);
                    Button bt = FindName($"box{l}{m}{n}") as Button;
                    if (sender is ComboBox)
                    {
                        ComboBox sl = FindName($"select{l}{m}{n}") as ComboBox;
                        ComboBoxItem typeItem = (ComboBoxItem)sl.SelectedItem;
                        if (typeItem == null)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else if (typeItem.Content == null)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else
                        {
                            string value = typeItem.Content.ToString();
                            if (value == "-Select-")
                            { bt.Background = new SolidColorBrush(Colors.Red); }
                            else
                            { bt.Background = new SolidColorBrush(Colors.Green); }
                        }
                    }
                    else if (sender is TextBox)
                    {
                        TextBox tx = FindName($"text{l}{m}{n}") as TextBox;
                        if (tx == null)
                        { MessageBox.Show("ERROR: textbox name is null"); }
                        else
                            if (tx.Text.Length == 0)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else
                        { bt.Background = new SolidColorBrush(Colors.Green); }
                    }
                    else if (sender is DatePicker)
                    {
                        bt.Background = new SolidColorBrush(Colors.Green);
                    }
                }
                else if (Char.IsNumber(selectName, selectName.Length - 2) == true)
                {
                    // two digit numbers
                    int m = (int)Char.GetNumericValue(selectName[selectName.Length - 2]);
                    Button bt = FindName($"box{m}{n}") as Button;
                    if (sender is ComboBox)
                    {
                        ComboBox sl = FindName($"select{m}{n}") as ComboBox;
                        ComboBoxItem typeItem = (ComboBoxItem)sl.SelectedItem;
                        if (typeItem == null)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else if (typeItem.Content == null)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else
                        {
                            string value = typeItem.Content.ToString();
                            if (value == "-Select-")
                            { bt.Background = new SolidColorBrush(Colors.Red); }
                            else
                            { bt.Background = new SolidColorBrush(Colors.Green); }
                        }
                    }
                    else if (sender is TextBox)
                    {
                        TextBox tx = FindName($"text{m}{n}") as TextBox;
                        if (tx == null)
                        { MessageBox.Show("ERROR: textbox name is null"); }
                        else
                            if (tx.Text.Length == 0)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else
                        { bt.Background = new SolidColorBrush(Colors.Green); }
                    }
                    else if (sender is DatePicker)
                    {
                        DatePicker dp = FindName($"date{m}{n}") as DatePicker;
                        if (dp.SelectedDate == null)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else
                        { bt.Background = new SolidColorBrush(Colors.Green); }
                    }
                }
                else
                {
                    // single digit numbers
                    // Will need to add another if for 4 digit number if user inputs go over 999!
                    Button bt = FindName($"box{n}") as Button;
                    if (sender is ComboBox)
                    {
                        ComboBox sl = FindName($"select{n}") as ComboBox;
                        ComboBoxItem typeItem = (ComboBoxItem)sl.SelectedItem;
                        if (typeItem == null)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else if (typeItem.Content == null)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else
                        {
                            string value = typeItem.Content.ToString();
                            if (value == "-Select-")
                            { bt.Background = new SolidColorBrush(Colors.Red); }
                            else
                            { bt.Background = new SolidColorBrush(Colors.Green); }
                        }
                    }
                    else if (sender is TextBox)
                    {
                        TextBox tx = FindName($"text{n}") as TextBox;
                        if (tx == null)
                        { MessageBox.Show("ERROR: textbox name is null"); }
                        else
                            if (tx.Text.Length == 0)
                        { bt.Background = new SolidColorBrush(Colors.Red); }
                        else
                        { bt.Background = new SolidColorBrush(Colors.Green); }
                    }
                    else if (sender is DatePicker)
                    {
                        bt.Background = new SolidColorBrush(Colors.Green);
                    }
                }
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MyScrollViewer.ScrollToVerticalOffset(530);
            buttonColor("ph");
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            MyScrollViewer.ScrollToVerticalOffset(0);
            buttonColor("rd");
        }
        private void MyScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (MyScrollViewer.VerticalOffset < 520)
            { buttonColor("rd"); }
            else
            { buttonColor("ph"); }
        }


        // User inputs
        private void select1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        { boxColor(sender); }
        private void select2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        { boxColor(sender); }
        private void text3_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text4_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text5_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text6_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text7_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void select8_SelectionChanged(object sender, SelectionChangedEventArgs e)
        { boxColor(sender); }
        private void text9_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text10_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text11_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text12_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text13_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void text14_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void date15_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        { boxColor(sender); }
        // number referes to row of item, there are no items in row 16, therfore this is missed out in numbering
        private void text17_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
        private void date18_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        { boxColor(sender); }
        private void text19_TextChanged(object sender, TextChangedEventArgs e)
        { boxColor(sender); }
    }
}
