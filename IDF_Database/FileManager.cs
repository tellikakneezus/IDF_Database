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
        Splash splash = new Splash(); // eventually find a way to move to MainWindow.

        string[][] dataIn; //
        string[,] cableInv;

        //Need to figure out how to get this to work
        //const string cableInvFilePath = @"CableInventory\Cables.xlsx";

         
        /**
         * Creation of FileManager object and Excel application and any files not processed yet will be processed. 
         * */
        public FileManager()
        {
            excel = new _Excel.Application(); //start excel

            updateFileManager();
            cleanExcelProcess();

        }

        public void updateFileManager()
        {
            splash.Show();
            splash.changeMessage("Getting Files"); string[] labelFiles = getFileNames("Labels");

            dataIn = processFiles(labelFiles);

            cleanExcelProcess();
            excel.Quit();
            splash.changeMessage("Moving files to Read"); filesRead(labelFiles);

            #region inventory logic. Deleting
            //splash.changeMessage("Opening Inventory"); openCableInventoryExcel();

            //splash.changeMessage("Importing Inventory"); cableInv = createCableInvArray();


            //splash.changeMessage("Updating Inventory"); updateCableInv(dataIn);
            #endregion


            splash.Hide();

            cleanExcelProcess();
        }


        /**
         * parses through all the files and puts the data into a staggered array. 
         * */
        private string[][] processFiles(string[] fileNames)
        {
            List<string[]> dataList = new List<string[]>(); //holds all data imported from label spreadsheets
            
            FileInfo fi;
            int count = 0;
            foreach (string file in fileNames)
            {
                count++;
                splash.changeMessage("Processing Files " + count + " out of " + fileNames.Length + ": " + file);
                fi = new FileInfo(file);
                openWorkbook(fi.FullName);
                
                ws = wb.ActiveSheet;
                range = ws.UsedRange;

                //read worksheet cell by cell and input data into dataList. 
                //i and j start at 1 because Cells do not start at 0,0 they start at 1,1 in Excel

                int numRows = range.Rows.Count;
                int numCols = range.Columns.Count;
                object[,] importArray = range.Cells.Value2;
                string[] objToStringLine;
                for (int i = 1; i <= numRows; i++)
                {
                    if(importArray[i,1] == null && importArray[i,2] == null)
                    {
                        break; //if nothing is in n column on row, go to next line
                    }
                    else
                    {
                        objToStringLine = new string[numCols];
                        for (int j = 0; j < numCols; j++)
                        {
                            if (importArray[i, j + 1] != null) objToStringLine[j] = importArray[i, j + 1].ToString(); else objToStringLine[j] = "";
                        }
                    }
                      
                    
                    dataList.Add(objToStringLine);
                }

                
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
                MessageBox.Show("Could not open " + filename + ". Please close make sure it is not open in another window." , "Important Message");
                
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
                MessageBox.Show("Error: File Not Found Exception while trying to get filenames from LABELS", "Important Message");
            }

            return fileNames;
        }

        /**
         * creates fileNames object. called from import data method. 
         * */
        private string[] getFileNames(string dir)
        {

            string[] fileNames = new string[0];
            try
            {
                fileNames = Directory.GetFiles(dir + "/");

            }
            catch (FileNotFoundException e)
            {
                MessageBox.Show("Error: File Not Found Exception while trying to get filenames from " + dir, "Important Message");
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
                    MessageBox.Show("Error trying to move file into read. Make sure the filename " + fileName + " is not being used by another file in READ folder.", "Important Message");
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
                    int divisions = 1;

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
                
                

            }
            if (!File.Exists(fileName))
            {
                wb.SaveAs(fileName);
                cleanExcelProcess();
                return true;
            }
            else
            {
                File.Delete(fileName);
                wb.SaveAs(fileName);
                cleanExcelProcess();
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
            string ppString = pp.ToLower().Substring(1);
            switch (ppString)
            {
                case "2s":
                    ppInt = 3;
                    break;
                case "2n":
                    ppInt = 4;
                    break;
                case "3s":
                    ppInt = 5;
                    break;
                case "3n":
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
        
        private string[,] formatPP(string[] pp, int height, int width)
        {
            string[,] fPP = new string[pp.Length / height, pp.Length / width];
            int count = 0;
            for (int i = 0; i < 36; i++)
            {
                for (int j = 0; j < 8; j++)
                {
                    fPP[i, j] = (count + 1).ToString() + "||" + pp[count];
                    count++;
                }
            }

            return fPP;
        }

        #region Cable Inventory Logic. Deleting
        //private string[,] createCableInvArray()
        //{

        //    //read worksheet cell by cell and input data into cableInv. 
        //    //i and j start at 1 because Cells do not start at 0,0 they start at 1,1 in Excel
        //    int numRows = range.Rows.Count;
        //    int numCols = range.Columns.Count;
        //    string[,] cableInv = new string[numRows, numCols];

        //    object[,] importArray = range.Cells.Value2;
        //    for (int i = 0; i <numRows; i++)
        //    {
        //        for (int j = 0; j < numCols; j++)
        //        {
        //            if (importArray[i+1, j + 1] != null) cableInv[i,j] = importArray[i+1, j + 1].ToString(); else cableInv[i, j] = "";
        //        }

        //    }
          

        //    return cableInv;

        //}


        ///**
        // * Opens the cable inventory WB and reassigns ws and range
        // * */
        //private void openCableInventoryExcel()
        //{
        //    string[] fileNames = getFileNames("CableInventory");

        //    //open first workbook in array
        //    FileInfo fi;
        //    fi = new FileInfo(fileNames[0]);
        //    openWorkbook(fi.FullName);

        //    //create worksheet and range
        //    ws = wb.ActiveSheet;
        //    range = ws.UsedRange;
        //}

        ///**
        // *  This will update Cable Inventory after data import. 
        // * */ 
        // private void updateCableInv(string[][] data)
        //{

        //    int[] cableInvCount = createCableInvCount();

        //    //check the first column of each data row for matching string in CableInv[0,i]. If true decriment count in cableInvCount.
        //    foreach(string[] row in data)
        //    {
        //        if(row != null)
        //        {
        //            for (int i = 0; i < cableInv.GetLength(0); i++)
        //            {
        //                if (string.Equals(row[0].ToLower(), cableInv[i, 0].ToLower()))
        //                {
        //                    cableInvCount[i]--;
        //                }
        //            }
        //        }
                
        //    }

        //    for (int i = 0; i < cableInv.GetLength(0); i++)
        //    {
        //        if (cableInvCount[i] > -1000) // if createCableInvCount() was unable to parse string to int it set value to -1000. 
        //        {
        //            cableInv[i, 1] = cableInvCount[i].ToString();
        //        }
        //    }

            

        //    for (int i = 0; i < cableInv.GetLength(0); i++)
        //    {
        //        if (cableInv[i,1] != null)
        //        {
        //            range.Cells[i+1, 2].Value2 = cableInv[i, 1];

        //        }
        //    }

        //    wb.Save();
            

        //}

        ///**
        // * Creates a one dimensional array of int values. It is parsed from string[,] cableInv. We use cableInvCount to update values on each cable length.
        // * */
        //private int[] createCableInvCount()
        //{
        //    int[] cableInvCount = new int[cableInv.GetLength(0)];
        //    for (int i = 0; i < cableInv.GetLength(0); i++)
        //    {
        //        try
        //        {
        //            cableInvCount[i] = Int32.Parse(cableInv[i, 1]);
        //        }
        //        catch (FormatException)
        //        {
        //            cableInvCount[i] = -1000; //best way i could think to handle it right now. Create a check when writing over CableInventory file for (< -79) and if true then place original value. Handles non numeric values in cableInv
        //        }
        //    }

        //    return cableInvCount;

        //}

        #endregion 

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
                MessageBox.Show("Error found while trying to create excel document", "Important Message");
            }
            finally
            {
            }

            return wb;

        }

        /**
         * Clean up methods
         * */

        private void cleanExcelProcess()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            try
            {
                while (Marshal.ReleaseComObject(range) > 0) ;
            }
            catch { }
            finally
            {
                range = null;
            }

            try
            {
                while (Marshal.ReleaseComObject(ws) > 0) ;
            }
            catch { }
            finally
            {
                ws = null;
            }

            try
            {
                wb.Close();
                while (Marshal.ReleaseComObject(wb) > 0) ;
            }
            catch { }
            finally
            {
                wb = null;
            }
        }

        public void closeFM()
        {

            excel.Quit();
            try
            {
                while (Marshal.ReleaseComObject(excel) > 0) ;
            }
            catch { }
            finally
            {
                excel = null;
            }
            
        }


    }
}