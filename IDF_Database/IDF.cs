using System;
using System.Runtime.Serialization;

namespace IDF_Database
{
    [Serializable()]
    internal class IDF : ISerializable
    {
        string name;
        string[] pps;
        int portsPerPP;
        string[][] data;

        public IDF(string nameIn, string[] ppsIn, int portsPerPPIn)
        {
            name = nameIn.ToLower();
            pps = ppsIn;
            foreach (string pp in pps)
            {
                pp.ToLower();
            }
            portsPerPP = portsPerPPIn;
            initailizeData(pps.Length); //data is null
        }

        public IDF(string nameIn, string[] ppsIn, int portsPerPPIn, string[][] dataIn)
        {
            name = nameIn.ToLower();
            pps = ppsIn;
            foreach (string pp in pps)
            {
                pp.ToLower();
            }
            portsPerPP = portsPerPPIn;
            data = dataIn;

        }

        private void initailizeData(int numPPs)
        {
            data = new string[numPPs][];
            for (int i = 0; i < data.Length; i++) //data is null
            {
                data[i] = new string[portsPerPP];
            }
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("name", name);
            info.AddValue("pps", pps);
            info.AddValue("portsPerPP", portsPerPP);
            info.AddValue("data", data);
        }

        public IDF(SerializationInfo info, StreamingContext context)
        {
            name = (string)info.GetValue("name", typeof(string));
            pps = (string[])info.GetValue("pps", typeof(string[]));
            portsPerPP = (int)info.GetValue("portsPerPP", typeof(int));
            data = (string[][])info.GetValue("data", typeof(string[][]));
        }

        public void insertData(string pp, int port, string info)
        {
            for (int i = 0; i < pps.Length; i++)
            {
                if (pp.ToLower() == pps[i])
                {
                    data[i][port - 1] = info;
                }
            }
        }

        public string getName()
        {
            return name;
        }

        public string[][] getData()
        {
            return data;
        }

        public string[] getPPs()
        {
            return pps;
        }
    }
}