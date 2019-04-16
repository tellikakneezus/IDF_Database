using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows;


namespace IDF_Database
{
    internal class FileManager
    {
        _Application excel;
        Workbook wb;
        Worksheet ws;
        Range range;

        string[][] dataIn; //
        string[,] cableInv;

        const string cableInvFilePath = "@SavedData/CableInventory/CableInventory.xlsx";

         
        /**
         * At creation of FileManager object and excel application is created and any files not processed yet will be processed. 
         * */
        public FileManager()
        {
            excel = new _Excel.Application();
            dataIn = importData();
        }

        /**
         * import data will take in every file in labels directory process them and then place them in the Read folder.
         * Processing includes: taking in any line in the label spreadsheet that contains a IDF cable and placing it in the data 2d string array
         * 
         *  
         * */
        private string[][] importData()
        {
            string[][] data = new string[0][];
            string[] fileNames = getFileNames();
            data = processFiles(fileNames); 

            filesRead(fileNames);
            return data;

        }


        /**
         * parses through all the files and puts the data into a staggered array. 
         * */
        private string[][] processFiles(string[] fileNames)
        {
            List<string[]> dataList = new List<string[]>();
            foreach(string file in fileNames)
            {

                FileInfo fi;
                fi = new FileInfo(file);
                openWorkbook(fi.FullName);
                ws = wb.ActiveSheet;
                range = ws.UsedRange;
                
                //read worksheet cell by cell and input data into dataList. 
                int numRows = range.Rows.Count;
                for (int i = 1; i <= numRows; i++)
                {
                    if (range.Cells[i, 2].Value2 != null) {
                        string[] line = new string[range.Columns.Count];
                        for (int j = 1; j <= line.Length; j++)
                        {
                            if (range.Cells[i, j].Value2 != null)
                            {
                                line[j - 1] = range.Cells[i, j].Value2.ToString();
                            }
                        }
                        dataList.Add(line);
                    }
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(range);
                
                Marshal.ReleaseComObject(ws);
                wb.Close();
                Marshal.ReleaseComObject(wb);

            }
            
           
            return dataList.ToArray(); 
        }

        /**
         * Created to handle the spreadsheet being open when user tries to start program
         * */
         private void openWorkbook(string filename)
        {
           
            try
            {
                wb = excel.Workbooks.Open(filename);
            
            }
            catch (COMException)
            {
                MessageBox.Show("Could not process files due to one or more being open. Please close all excel documents and restart program.", "Important Message");
                
            }
           
        }

        /**
         * creates fileNames object. called from import data method
         * */
        private string[] getFileNames()
        {
           
            string[] fileNames = new string[0];
            try
            {
                fileNames = Directory.GetFiles("Labels/");
                
            }
            catch (FileNotFoundException e)
            {
                Console.Write(e);
            }

            return fileNames;
        }

        /**
         * used to handle the moving of files into read folder.
         * */
        private void filesRead(string[] fileNames)
        {
            foreach (string fileName in fileNames)
            {
                string[] pathAndName = fileName.Split('/'); //for relocating file
                try
                {
                    File.Move(fileName, (pathAndName[0] + "/Read/" + pathAndName[1]));
                }
                catch (IOException)
                {
                    MessageBox.Show("Could not process files due to one or more being open. Please close all excel documents and restart program.", "Important Message");
                }

            }

        }


        /**
         * Getter for dataIn. Used in Main to pass on imported data to IDF_Db class
         * */
        public string[][] getDataIn()
        {
            return dataIn;
        }

        
        /**
         * Write Excel WB will take in:
         * -File name and location it will save exported document in
         * -Names of IDFs in IDF_DB || string[]
         * -PPs that go to corresponding IDF_names ||string[][]
         * -All data in IDF_DB || string[][][]
         * 
         * Will write excel workbook at location provided and return true if it writes the document successfully. 
         * */

        public bool writeIdfDbExcelWB(string fileName, string[] allIdfNames, string[][] allIdfPPs, string[][][] allIdfData)
        {

            wb = createExcelDoc(allIdfNames.Length); //this creates a workbook with the same number of worksheets as there are IDF cabinets
            int startX, startY, endX, endY; // For range editing.
            for(int i = 0; i< allIdfNames.Length; i++) //loop through each name in IDF name list
            {
                
                
                if (wb.Worksheets[i+1] != null)ws = wb.Worksheets[i+1]; //if a worksheet was created at the index (should always be true if createExcelDoc() did its job
                else { ws = new Worksheet(); wb.Worksheets.Add(After: ws); } //if a worksheet was not created it creates one and adds to end of worksheet list
                ws.Name = allIdfNames[i]; //creates name for worksheet


                startX = startY = endY = 1; //starts range at top left of worksheet
                endX = 36; //range will be [1,1] -> [36,1]

                

                //parse through all idf pps
                for (int j = 0; j < allIdfPPs[i].Length; j++)
                {
                    //restart at the top of the page.
                    startX = 1;



                    //add title to PP data
                    range = (Range)ws.Range[ws.Cells[startX, startY], ws.Cells[endX, endY]]; 
                    range.Merge(); range.HorizontalAlignment = XlHAlign.xlHAlignCenter; range.Interior.Color = highlightColor(ppHighlightColorInt(allIdfNames[i],allIdfPPs[i][j])); //Merge,Center, Highlight
                    range.Value = allIdfPPs[i][j];

                    //change range to hold PP Data
                    startY++; endY += 8;

                    //number of divisions per PP
                    int divisions = 4;

                    //Reformat pp data and add to spreadsheet
                    string[,] inputData = formatPP(allIdfData[i][j]);
                    for(int k = 1; k <= divisions; k++) //separates info in sections for highlighting
                    {

                        string[,] temp = array2dSection(inputData,divisions,k);
                        
                        endX *= k;
                        endX /= divisions;
                        range = (Range)ws.Range[ws.Cells[startX, startY], ws.Cells[endX, endY]];
                        range.Interior.Color = highlightColor(k);
                        range.Value = temp;
                        startX = endX + 1;
                        endX = 36;
                        
                    }


                    //Drop 2 lines and change range to one row
                    startY += 10;//pass the endY then add 2 lines
                    endY = startY;

                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(range);

                Marshal.ReleaseComObject(ws);

            }
            if (!File.Exists(fileName))
            {
                wb.SaveAs(fileName);
                wb.Close();
                Marshal.ReleaseComObject(wb);
                return true;
            }
            else
            {
                File.Delete(fileName);
                wb.SaveAs(fileName);
                wb.Close();
                Marshal.ReleaseComObject(wb);
                return true;
            }
           
        }

        /**
        * Gives an integer value showing PP to IDF relationship. Integer will be used in the highlightColor() method to highlight the PP.
        * Takes in two strings and converts them into an integer and then adds them together. 
        * */
        private int ppHighlightColorInt(string idf, string pp)
        {
            int a = ppStringToInt(idf);
            int b = ppStringToInt(pp);
            int offset = 2;
            return (a + b) - offset;

        }

        /**
         * Converts room number into an integer. Used in ppHighlightColorInt() method
         * */
        private int ppStringToInt(string pp)
        {
            int ppInt = 0;

            switch (pp.ToLower())
            {
                case "02s":
                    ppInt = 3;
                    break;
                case "02n":
                    ppInt = 4;
                    break;
                case "03s":
                    ppInt = 5;
                    break;
                case "03n":
                    ppInt = 7;
                    break;

            }

            return ppInt;

        }

        /**
         * will return color based on which section of patch panel you are on. 
         * */
        private Color highlightColor(int k)
        {
            Color col = Color.White;

            switch (k)
            {
                case 1:
                    col = Color.LightBlue;
                    break;
                case 2:
                    col = Color.Aquamarine;
                    
                    break;
                case 3:
                    col = Color.LightGreen;
                    
                    break;
                case 4:
                    col = Color.MediumAquamarine;
                    break;
                case 5:
                    col = Color.LightGray;
                    break;
                case 6:
                    col = Color.MistyRose;
                    break;
                case 7:
                    col = Color.Plum;
                    break;
                case 8:
                    col = Color.Tomato;
                    break;
                case 9:
                    col = Color.LightPink;
                    break;
                case 10:
                    col = Color.LightSteelBlue;
                    break;
                default:
                    col = Color.White;
                    break;
            }

            return col;
        }

        /**
         * Slices 2d array into sections. This is utilized in the highlighting of different sections of a patch panel
         * */
        private string[,] array2dSection(string[,] inputData, int numCuts, int cut)
        {
            string[,] a2dSection = null;
            if(inputData.GetLength(0)%numCuts == 0) // handles if pp cannot be divided evenly by number of cuts.
            {
                a2dSection = new string[inputData.GetLength(0) / numCuts, inputData.GetLength(1)];
                for(int i = 0; i < a2dSection.GetLength(0); i++)
                {
                    int inpDatXPos = inputData.GetLength(0);
                    float divInpDatXPosFl = cut - 1;
                    divInpDatXPosFl /= numCuts;

                    divInpDatXPosFl *= inpDatXPos;
                    inpDatXPos = (int)divInpDatXPosFl;
                    inpDatXPos += i;

                     for(int j = 0; j < a2dSection.GetLength(1); j++)
                    {
                        a2dSection[i, j] = inputData[inpDatXPos, j];
                    } 
                }   
            }
            return a2dSection;
        }
        /**
         * puts all data into a 2d array that represents the patch panel. also adds port number in front of data. 
         * right now has a default size of [8,36]. **Future: will handle different size pps by giving the closest size ratio of [1,4]
         * */
        private string[,] formatPP(string[] pp)
        {
            string[,] fPP = new string[pp.Length / 8, pp.Length / 36];
            int count = 0;
            for (int i = 0; i < 36; i++)
            {
                for(int j = 0; j < 8; j++)
                {
                    fPP[i, j] = (count+1).ToString() + "||" + pp[count];
                    count++;
                }
            }

            return fPP;
        }
        

        /**
         * Creates a workbook the number of sheets passed in by parameter.
         * */
        public Workbook createExcelDoc(int size)
        {
            Workbook wb = null;

            try
            {
                excel.SheetsInNewWorkbook = size;
                wb = excel.Workbooks.Add();

            }
            catch (Exception)
            {
                Console.Write("Error");
            }
            finally
            {
            }

            return wb;

        }

        /**
         * Clean up process
         * */

        public void closeFM()
        {
            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }


    }
}