using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace IDF_Database
{
    internal class IDF_db
    {
        List<IDF> idfs;
        const string SaveFileName = @"SavedData/IDFs.bin";
        string fallBackFileName = @"SavedData/" + DateTime.Now.ToString("_MM-dd-yyyy_HHmm") + ".bin";

        Stream stream;
        BinaryFormatter bf;

        public IDF_db()
        {

            idfs = new List<IDF>();

            try
            {
                stream = File.Open(SaveFileName, FileMode.Open);
                bf = new BinaryFormatter();
                if(stream.Length > 0)
                {
                    idfs = (List<IDF>)bf.Deserialize(stream);
                    //Console.WriteLine(idfs); //used for debugging saving process
                }
                else
                {
                    defaultIDFs();
                }
                
                stream.Close();
            }
            catch(FileNotFoundException e)
            {
                Console.WriteLine(SaveFileName + " Not Found!");
            }
            
            
        }

        private void defaultIDFs()
        {
            addIDF("02n", new string[3] { "03n", "03s", "02s" }, 288);
            addIDF("02s", new string[3] { "03n", "03s", "02n" }, 288);
            addIDF("03s", new string[3] { "03n", "02n", "02s" }, 288);
            addIDF("03n", new string[3] { "03s", "02n", "02s" }, 288);
        }

        public void addIDF(string name, string[] pps, int portsPerPP)
        {
            idfs.Add(new IDF(name, pps, portsPerPP));
        }

        public void addIDF(IDF idf)
        {
            idfs.Add(idf);
        }

        public void changeFallbackName(string name)
        {
            fallBackFileName = @"SavedData/" + name + ".bin";
        }

        public void saveIDFs()
        {
            
            stream = File.Open(SaveFileName, FileMode.Create);
            bf = new BinaryFormatter();
            bf.Serialize(stream, idfs);
            stream.Close();

            stream = File.Open(fallBackFileName, FileMode.Create);
            bf = new BinaryFormatter();
            bf.Serialize(stream, idfs);
            stream.Close();
        }

        public void insertData(string[][] data)
        {
            foreach(string[] info in data)
            {
                if(info != null)
                {
                    int i = isIdfCable(info[1]);

                    if (i != -1)
                    {
                        string idf = info[1].Substring(0, 3).ToLower();
                        string ppInIDF = info[1].Substring(3, 3).ToLower();
                        string infoOut;
                        if (info[2] != "") //user is not trying to blank out port. 

                        {
                            infoOut = info[3] + ", " + info[2] + " | " + info[5] + ", " + info[4];
                        }
                        else
                        {
                            infoOut = "";
                        }
                        
                        int portInPP;
                        if (int.TryParse(info[1].Substring(6), out portInPP)) // returns false if info[1] starting at char6 cannot be converted to int and will not try to insert data. 
                        {
                            idfs[i].insertData(ppInIDF, portInPP, infoOut);
                            //need to update complimenting pp.
                            idfs[getIdfIndex(ppInIDF)].insertData(idf, portInPP, infoOut);
                        }
                    }
                }
               
            }
           
        }

        private int isIdfCable(string cable) //returns -1 if an IDF does not claim the name called
        {
            if(cable.Length > 6)
            {
                if (getIdfIndex(cable.Substring(3, 3)) != -1)
                {
                    return getIdfIndex(cable.Substring(0, 3));
                } 
                
            }
            return -1;
        }

        private int getIdfIndex(string name)
        {
            int index = -1; 
            for(int i = 0; i< idfs.Count; i++)
            {
                if (name.ToLower() == idfs[i].getName())
                {
                    return i;
                }
            }

            return index;
        }

        public string[][][] getAllIdfData()
        {
            
            List<string[][]> allIdfData = new List<string[][]>();
            foreach(IDF idf in idfs)
            {
                allIdfData.Add(idf.getData());
            }

            return allIdfData.ToArray();

        }

        public string[][] getAllIdfPPs()
        {
            List<string[]> allIdfPPs = new List<string[]>();
            foreach(IDF idf in idfs)
            {
                allIdfPPs.Add(idf.getPPs());
            }
            return allIdfPPs.ToArray();
        }

        public string[] getAllIdfNames()
        {
            List<string> allIdfNames = new List<string>();
            foreach(IDF idf in idfs)
            {
                allIdfNames.Add(idf.getName());
            }

            return allIdfNames.ToArray();
        }

        

        public bool fallBackData(string file)
        {

            string fileName = "SavedData/" + file + ".bin";
            try
            {
                stream = File.Open(fileName, FileMode.Open);
                bf = new BinaryFormatter();
                if (stream.Length > 0)
                {
                    idfs = (List<IDF>)bf.Deserialize(stream);
                    //Console.WriteLine(idfs); //used for debugging saving process
                }
                else
                {
                    defaultIDFs();
                }

                stream.Close();
                return true;

            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(SaveFileName + " Not Found!");
                return false;
            }
            
        }

    }
}