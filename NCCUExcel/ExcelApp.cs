using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.OleDb;
using System.Data.SqlClient;
using Excel = NetOffice.ExcelApi;
using System.Data.Common;
using System.Text.RegularExpressions;
using System.IO;

namespace NCCUExcel
{
    public class ExcelApp
    {
        public static string _TopFolder = "C:\\Users\\jim\\Desktop\\thesis\\1226\\";

        public static string _FileFolder = _TopFolder;
        public static string[] _RecordString = new string[] { "facebook", "line", "youtube", "gmail", "chrome", "googlemap"};

        public static void exportApp(string fileInputName, string fileOutputName)
        {
            List<DataStruct> datas = GetExcel(_FileFolder + fileInputName);
            List<DeviceRecordStruct> allDeviceTimes = new List<DeviceRecordStruct>();

            for (int i = 0; i < datas.Count; i++)
            {
                for(int j = 0 ; j < _RecordString.Count() ; j++)
                {
                    if (datas[i].Value.Equals(_RecordString[j]))
                    {
                        if (allDeviceTimes.Find(ById(datas[i].Id)) == null)
                        {
                            DeviceRecordStruct deviceRecord = new DeviceRecordStruct(datas[i].Id);
                            allDeviceTimes.Add(deviceRecord);
                        }

                        DeviceRecordStruct findDeviceRecord = allDeviceTimes.Find(ById(datas[i].Id));

                        findDeviceRecord.AllTimes[j]++;
                        break;
                    }
                }
            }


            using (System.IO.StreamWriter file = new System.IO.StreamWriter(_FileFolder + fileOutputName))
            {
                file.Write(",");

                for (int i = 0; i < allDeviceTimes.Count; i++)
                {
                    string ID = allDeviceTimes[i].Id + ",";
                    file.Write(ID);
                }

                file.WriteLine();

                for (int i = 0; i < _RecordString.Count(); i++)
                {
                    string show = _RecordString[i].ToString();

                    for (int j = 0; j < allDeviceTimes.Count(); j++)
                    {
                        show += "," + allDeviceTimes[j].AllTimes[i].ToString();
                    }

                    file.WriteLine(show);
                }
                    

            }
        }

        private static List<DataStruct> GetExcel(string path)
        {
            List<DataStruct> list = new List<DataStruct>();
            using (StreamReader sr = new StreamReader(path))
            {
                String line = sr.ReadToEnd();
                string[] records = line.Split('\n');

                foreach (var record in records)
                {
                    string[] row = record.Split(',');
                    DataStruct one = new DataStruct()
                    {
                        Id = Convert.ToInt16(row[0]),
                        Value = row[1].ToString(),
                        date = Convert.ToDateTime(row[2].Replace("\r", ":00"))
                    };

                    list.Add(one);
                }

            }

            return list;
        }

        private class DataStruct
        {
            public int Id { get; set; }
            public string Value { get; set; }
            public DateTime date { get; set; }
        }

        static Predicate<DeviceRecordStruct> ById(int id)
        {
            return delegate(DeviceRecordStruct deviceRecord)
            {
                return deviceRecord.Id == id;
            };
        }

        private class DeviceRecordStruct
        {
            public int Id { get; set; }
            public int[] AllTimes { get; set; }

            public DeviceRecordStruct(int id)
            {
                this.Id = id;
                AllTimes = new int[_RecordString.Count()];
            }
        }

    }


}