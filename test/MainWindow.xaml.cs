using System;
using Microsoft.Win32;
using System.Collections.Generic;
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
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Specialized;

namespace test
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

        // Important Buttons
        private void packTxt(string inputValuesFileName)
        {
            // Write all user inputs into a txt file
            int numberOfProperties = 1;
            while (FindName($"name{numberOfProperties}") != null)
            { numberOfProperties = numberOfProperties + 1; }
            numberOfProperties = numberOfProperties - 1;


            string[] first = new string[numberOfProperties + 1]; // string containing perameter names (defined by label content)
            string[] second = new string[numberOfProperties + 1]; // string containing values inputed by user (defined by textbox/ combobox)
            int i = 1;
            //Label test = FindName($"name{1}") as Label;
            //while (FindName($"name{i}") != null)
            while (i <= numberOfProperties)
            {
                Label lb = FindName($"name{i}") as Label;
                first[i] = (string)lb.Content;
                object isTextBox = FindName($"text{i}");
                object isComboBox = FindName($"select{i}");
                object isDatePicker = FindName($"date{i}");
                if (isTextBox != null)
                {
                    TextBox tb = FindName($"text{i}") as TextBox;
                    if (tb.Text != "")
                        second[i] = tb.Text;
                    else { second[i] = "-empty-"; }
                }
                else if (isComboBox != null)
                {
                    ComboBox cb = FindName($"select{i}") as ComboBox;
                    if (cb.Text != "")
                        second[i] = cb.Text;
                    else { second[i] = "-empty-"; }
                }
                else if (isDatePicker != null)
                {
                    DatePicker dp = FindName($"date{i}") as DatePicker;
                    if (dp.SelectedDate.ToString() != "")
                        second[i] = dp.SelectedDate.ToString();
                    else { second[i] = "-empty-"; }
                }
                i++;
            }

            //// Write first Column
            //var first = new[] { "PARAMETERS", ".--------Report Details---------", "Project---------------------------", "select1", "Reason for Div", "Project Name", "Infr Client/s", "Pipeline Name", "Pipeline Number", "Diversion", "Bore Size" };
            File.WriteAllLines(inputValuesFileName, first);

            // Write second Column and add it onto the fist in the txt file
            //var second = new[] { "VALUE", "", "", select1.Text, select2.Text, text3.Text, text4.Text, text5.Text, text6.Text, text7.Text, select8.Text };
            var file = File.ReadAllLines(inputValuesFileName);
            for (int ii = 0; ii < file.Length; ii++)
                file[ii] += '\t' + second[ii]; //add the second column after the first, with a tab
            File.WriteAllLines(inputValuesFileName, file);
        }
        private void unpackTxt(string inputValuesFileName)
        {
            // read input values file
            string phrase = System.IO.File.ReadAllText(inputValuesFileName);

            // split up the string into an array and delete un-necessary rows
            string[] words = phrase.Split('\t', '\n', '\r');
            string[] newWords = new string[words.Length];
            var ii = 0;
            for (int i = 0; i < words.Length; i++)
                // If that string doesnt contain '-' and it isnt length 0 then it will be included in the new list
                // if (words[i].Contains('-') == false && words[i].Length != 0)
                if (words[i].Length != 0)
                {
                    newWords[ii] = words[i];
                    ii = ii + 1;
                }

            for (int i = 0; i < newWords.Length; i=i+2)
            {
                int n = (i + 2) / 2;
                if (FindName($"select{n}") != null)
                {
                    ComboBox tb = FindName($"select{n}") as ComboBox;
                    tb.Text = newWords[i + 1];
                }
                else if (FindName($"text{n}") != null)
                {
                    TextBox tb = FindName($"text{n}") as TextBox;
                    if (newWords[i + 1] != "-empty-")
                    { tb.Text = newWords[i + 1]; }
                    else
                    { tb.Text = ""; }
                }
                else if (FindName($"date{n}") != null)
                {
                    DatePicker dp = FindName($"date{n}") as DatePicker;
                    if (newWords[i + 1] != "-empty-")
                    {dp.SelectedDate = Convert.ToDateTime(newWords[i + 1]); } // add code to say 'select a date' is newWords[i + 1] = "-empty-"
                }
            }

            //TextBox tb = FindName("select1") as TextBox;
            //tb.Text = newWords[1];
            //select2.Text = newWords[3];
            //text3.Text = newWords[5];
            //text4.Text = newWords[7];
            //text5.Text = newWords[8];
            //text6.Text = newWords[10];
            //text7.Text = newWords[12];
            //select8.Text = newWords[14];
        }
        private void temporyFileNameStore(string txtFileName)
        {
            //  Stores txtFileName in a new txt file (saveasFileLoaction.txt)

            //  allows txtFileName to be brought back from saveasFileLocation later
            //  (this new txt file is deleted when the application is closed)

            var saveasFileLocation = @"C:\Users\sevantej\OneDrive - Jacobs\Documents\Technical\Report Automation\AutomatedDocuments\saveasFileLoaction.txt";
            //if (File.Exists(saveasFileLocation)) { File.Delete(saveasFileLocation); }
            File.WriteAllText(saveasFileLocation, txtFileName);
        }
        private void emmasSaveas()
        {
            // Open save as dialog and describe the dialog options (i,e. save as a txt file, start in AutomatedDocument folder, etc)
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = @"C:\Users\sevantej\OneDrive - Jacobs\Documents\Technical\Report Automation\AutomatedDocuments\";
            saveFileDialog.Filter = "Text Files | *.txt";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;

            // if user clicks ok/save
            if (saveFileDialog.ShowDialog() == true)
            {
                // Create txt file using inputValuesTxt function (defined above)
                packTxt(saveFileDialog.FileName);
                MessageBox.Show("Options Saved-as!");
                
                // Call temporyFileNameStore function to store the file name so it can be retrieved later
                temporyFileNameStore(saveFileDialog.FileName);
            }
            // elseif user didn't click ok/save it will display the message below
            else { MessageBox.Show("Could not save"); }
        }
        private void emmasSave()
        {
            try
            {
                // Check if file already exists. If yes, delete it.     
                var saveasFileLocation = @"C:\Users\sevantej\OneDrive - Jacobs\Documents\Technical\Report Automation\AutomatedDocuments\saveasFileLoaction.txt";
                if (File.Exists(saveasFileLocation))
                {
                    string fileName = System.IO.File.ReadAllText(saveasFileLocation);
                    packTxt(fileName);
                    MessageBox.Show("Options Saved!");
                }
                else { emmasSaveas(); }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.ToString());
            }

        }
        private void save_Click(object sender, RoutedEventArgs e)
        {  emmasSave(); }
        private void saveas_Click(object sender, RoutedEventArgs e)
        {
            emmasSaveas();
        }
        private void Open_Click(object sender, RoutedEventArgs e)
        {
            // need to prompt user to open file then save file location to saveasFileLocation txt file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                // Call unpackTxt function to read txt file and add values to the 
                unpackTxt(openFileDialog.FileName);
                // Call temporyFileNameStore function to store the opened txt's file name
                temporyFileNameStore(openFileDialog.FileName);
            }
        }
        static public void ReleaseComObjectsUsingGC()
        {
            // Help from Brian Boye: https://github.com/People-Places-Solutions/Geo-Digital_ExcelWrapper 
            
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
        private void goButton_Click(object sender, RoutedEventArgs e)
        {
            emmasSave();
            MessageBox.Show("This may take a minute, message box will pop up when done");
            /* Help from Brian Boye:
            https://github.com/People-Places-Solutions/Geo-Digital_ExcelWrapper */
            try { updateDoc();}
            finally {ReleaseComObjectsUsingGC();}

            void updateDoc()
            {
                // Define filePath, where the document CDS_PROPS is saved
                String filePath = @"C:\Users\sevantej\OneDrive - Jacobs\Documents\Technical\Report Automation\AutomatedDocuments\CDS_PROPS.docx";

                try
                {
                    // Make sure file exists
                    if (File.Exists(filePath))
                    {
                        // Code from: https://social.msdn.microsoft.com/Forums/sqlserver/en-US/8dc4afdf-8d12-4b6e-8de8-fc990f39c8c1/creating-n-accessing-custombuiltin-document-properties?forum=worddev

                        // Define perameters for later
                        object missing = Missing.Value;
                        object DocCustomProps;

                        // Define aDoc as document and wordApp as the word application
                        Word.Application wordApp = null;
                        Word._Document aDoc = null;
                        wordApp = new Word.Application();
                        
                        // Open and activate word doc containing custom propeties
                        aDoc = wordApp.Documents.Open(filePath, ReadOnly: false);
                        aDoc = wordApp.ActiveDocument;

                        //Get the CustomDocumentProperties collection and find out type.
                        DocCustomProps = aDoc.CustomDocumentProperties;
                        Type typeDocCustomProps = DocCustomProps.GetType();

                        string strIndex = "Client Name";
                        string strValue = select1.Text;

                        typeDocCustomProps.InvokeMember("Item",
                                                   BindingFlags.Default |
                                                   BindingFlags.SetProperty,
                                                   null, DocCustomProps,
                                                   new object[] { strIndex, strValue });

                        strIndex = "Infrastructure Project Name";
                        strValue = text3.Text;

                        typeDocCustomProps.InvokeMember("Item",
                                                   BindingFlags.Default |
                                                   BindingFlags.SetProperty,
                                                   null, DocCustomProps,
                                                   new object[] { strIndex, strValue });

                        strIndex = "Infrastructure Long Name";
                        strValue = text4.Text;

                        typeDocCustomProps.InvokeMember("Item",
                                                   BindingFlags.Default |
                                                   BindingFlags.SetProperty,
                                                   null, DocCustomProps,
                                                   new object[] { strIndex, strValue });


                        // Update Property Fields in Document
                        aDoc.Fields.Update();

                        //Save the document  
                        aDoc.Save();
                        MessageBox.Show("File Saved!");

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
            //{ updateEx(); }

            void updateEx()
            {
                // Define filePath, where the document CDS_PROPS is saved
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
        }
        private void Window_Closed(object sender, EventArgs e)
        {
            var saveasFileLocation = @"C:\Users\sevantej\OneDrive - Jacobs\Documents\Technical\Report Automation\AutomatedDocuments\saveasFileLoaction.txt";
            File.Delete(saveasFileLocation);
        }

        // Scroller/ placer buttons
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
                        bt.Background = new SolidColorBrush(Colors.Green);
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
            if (MyScrollViewer.VerticalOffset < 530)
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
    }
}
