using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using Timer = System.Timers.Timer;
using System.ComponentModel;


namespace IDF_Database
{

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
    public partial class MainWindow : Window
    {

        FileManager fm;
        IDF_db idfDB;

        enum TutorialStep { Start, Fallback, Spreadsheet, End }; //used to track where user is in tutorial.
        bool tutorialFlag = true; //set when user chooses tutorial option. 


     
        

        public MainWindow()
        {

            initializeFM_IdfDB();
            InitializeComponent();
            enableMenu();
            Tutorial(TutorialStep.Start);                   
        }


        private void MainForm_Closing(object sender, CancelEventArgs e)
        {
            idfDB.saveIDFs();
            fm.closeFM();
        }

        private void SaveIdfFile_Click(object sender, RoutedEventArgs e)
        {
            fallbackButton.Visibility = Visibility.Hidden;
            fallbackListBox.Visibility = Visibility.Hidden;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel |*.xlsx";
            if(sfd.ShowDialog() == true)
            {
                
                if(fm.writeIdfDbExcelWB(sfd.FileName,idfDB.getAllIdfNames(), idfDB.getAllIdfPPs(), idfDB.getAllIdfData()))
                {
                    Tutorial(TutorialStep.Spreadsheet);
                }
                
            }
        }

        private void initializeFM_IdfDB()
        {
            fm = new FileManager();
            idfDB = new IDF_db();
            idfDB.insertData(fm.getDataIn());


        }

        private void enableMenu()
        {
            FileMenu.IsEnabled = true;
            AddMenu.IsEnabled = true;
            HelpMenu.IsEnabled = true;
        }

        private void FallbackDatabase_Click(object sender, RoutedEventArgs e)
        {

            fallbackButton.Visibility = Visibility.Visible;
            fallbackListBox.Visibility = Visibility.Visible;
            Tutorial(TutorialStep.Fallback);

            string[] fileNames = new string[0];
            try
            {
                fileNames = Directory
                .EnumerateFiles("SavedData/", "*.bin", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileNameWithoutExtension)
                .ToArray(); //list all the file names in the SavedData/ path without extension and ignore any files in subdirectories. ||FUTURE: Get this list from File Manager.
                    

            }
            catch (FileNotFoundException fnfE)
            {
                Console.Write(fnfE);
            }

            foreach(string file in fileNames)
            {
                if(file != "IDFs")
                {
                    fallbackListBox.Items.Add(file);
                }
            }

        }

        private void SaveEventButton_Click(object sender, RoutedEventArgs e)
        {
            idfDB.changeFallbackName(custFallbackNameTextBox.Text);
            custFallbackNameTextBox.Text = "Success";
        }

        private void FallbackButton_Click(object sender, RoutedEventArgs e)
        {
            Timer timer = new Timer();
            

            string file = fallbackListBox.SelectedItem.ToString();
            if (file != null) //if nothing is selected do nothing.
            {
                if (idfDB.fallBackData(file))
                {
                    //update status label with successful update
                    statusLabel.Content = "Fallback to " + file + " successful";
                    //hide UI elements
                    fallbackButton.Visibility = Visibility.Hidden;
                    fallbackListBox.Visibility = Visibility.Hidden;

                    //set timer for 5 seconds and hide text in status label
                    timer = new Timer();
                    timer.Interval = 5000;
                    timer.Elapsed += OnTimedEvent;
                    timer.AutoReset = true;
                    timer.Enabled = true;

                }
                else
                {
                    //update status label with unsuccessful update
                    statusLabel.Content = "Fallback to " + file + " unsuccessful";

                    //set timer for 5 seconds and hide text in status label
                    timer = new Timer();
                    timer.Interval = 5000;
                    timer.Elapsed += OnTimedEvent;
                    timer.AutoReset = true;
                    timer.Enabled = true;
                }
                
            }
            
        }

        private void AddIdf_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
        {
            this.Dispatcher.Invoke(() =>
            {
                statusLabel.Content = "";
            });
            
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            nextButton.Visibility = Visibility.Hidden;
            Tutorial(TutorialStep.End);
        }

        private void Tutorial(TutorialStep step)
        {


            if (tutorialFlag)
            {
                switch (step)
                {
                    case TutorialStep.Start:
                        tutLabel.Visibility = Visibility.Visible;
                        tutLabel.Content = "Welcome to IDF Database. \n\nIf you already know where open ports are: \nCreate a label spreadsheet with the 'N' Column showing the ports that you chose.\n" +
                            "An example of a port would be '02n02s38'.\n\n Go ahead and place your label spreadsheet in the labels folder of this program and then close out \n\nIf you do not know where open ports are you can save a spreadsheet of all ports by clicking: File-> Save IDF File" +
                            "\n\nNotice the text box on the bottom of the screen. This text box is used to create a custom fallback marker for the future.";
                        nextButton.Visibility = Visibility.Hidden;
                        break;
                    case TutorialStep.Spreadsheet:
                        tutLabel.Content = "Now that the file is saved. Open it and find empty ports. You can tell if the port is empty if there is no information on it. \n" +
                            "Create a label spreadsheet with the 'N' Column showing the ports that you chose.\n\n" +
                            "An example of a port would be '02n02s38'.\n" +
                            "'02n' is the IDF cabinet '02s' is the PP in the cabinet going to 2s and '38' is the port number on the patch panel.\n\n" +
                            "The program automatically updates both ports for the cable going through the IDF cabinets.";
                        nextButton.Visibility = Visibility.Visible;
                        break;
                    case TutorialStep.Fallback:
                        tutLabel.Visibility = Visibility.Hidden;
                        nextButton.Visibility = Visibility.Hidden;
                        
                        break;
                    case TutorialStep.End:
                        tutLabel.Content = "Save the label spreadsheet in the Labels folder of the program and close the program.\n" +
                            "You are done. The database will be updated the next time the program is used.\n\n" +
                            "If you used bad data. There will be a save point created everytime the program is closed.\n " +
                            "Just go to File-> Fallback Database and then choose the date that you would like to fallback to.";
                        break;

                }
            }
            
        }

        private void Tutorial_Click(object sender, RoutedEventArgs e)
        {
            tutorialFlag = true;
            Tutorial(TutorialStep.Start);
        }

        private void ImportLabels_Click(object sender, RoutedEventArgs e)
        {
            fm.updateFileManager();
        }
    }
}
