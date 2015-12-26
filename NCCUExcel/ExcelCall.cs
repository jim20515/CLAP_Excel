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
    public class ExcelCall
    {
        public static string _TopFolder = "C:\\Users\\jim\\Desktop\\thesis\\1226\\";

        public static string _FileFolder = _TopFolder;

        public static void exportCall(string dataName, string fileInputName, string fileOutputName)
        {
            List<DataStruct> datas = GetExcel(_FileFolder + dataName);
            List<DeviceRecordStruct> allDeviceTimes = new List<DeviceRecordStruct>();

            for (int i = 0; i < datas.Count; i++)
            {
                if (allDeviceTimes.Find(ById(datas[i].Id)) == null)
                {
                    DeviceRecordStruct deviceRecord = new DeviceRecordStruct(datas[i].Id);
                    allDeviceTimes.Add(deviceRecord);
                }

                DeviceRecordStruct findDeviceRecord = allDeviceTimes.Find(ById(datas[i].Id));

                if (datas[i].isOut)
                {
                    findDeviceRecord.AllOutTimes[datas[i].date.Hour] += datas[i].Value;
                }
                else
                {
                    findDeviceRecord.AllInTimes[datas[i].date.Hour] += datas[i].Value;
                }
            }


            using (System.IO.StreamWriter file = new System.IO.StreamWriter(_FileFolder + fileOutputName))
            {
                file.Write(" ,");

                for (int i = 0; i < allDeviceTimes.Count; i++)
                {
                    string ID = allDeviceTimes[i].Id + ",";
                    file.Write(ID);
                }

                file.WriteLine();

                for (int i = 0; i < 24; i++)
                {
                    string show = i.ToString();

                    for (int j = 0; j < allDeviceTimes.Count(); j++)
                    {
                        show += "," + allDeviceTimes[j].AllOutTimes[i].ToString();
                    }

                    file.WriteLine(show);
                }
            }
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(_FileFolder + fileInputName))
            {
                file.Write(" ,");

                for (int i = 0; i < allDeviceTimes.Count; i++)
                {
                    string ID = allDeviceTimes[i].Id + ",";
                    file.Write(ID);
                }

                file.WriteLine();

                for (int i = 0; i < 24; i++)
                {
                    string show = i.ToString();

                    for (int j = 0; j < allDeviceTimes.Count(); j++)
                    {
                        show += "," + allDeviceTimes[j].AllInTimes[i].ToString();
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
                    //jim
                    string formatRecord = record.Replace("\r", "");
                    string[] row = formatRecord.Split('\t');
                    DataStruct one = new DataStruct()
                    {
                        Id = Convert.ToInt32(row[0]),
                        isOut = "撥出".Equals(row[1]),
                        Value = Convert.ToInt32(row[3]),
                        date = Convert.ToDateTime(row[4])
                    };

                    list.Add(one);
                }

            }

            return list;
        }

        private class DataStruct
        {
            public int Id { get; set; }
            public bool isOut { get; set; }
            public int Value { get; set; }
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
            public int[] AllInTimes { get; set; }
            public int[] AllOutTimes { get; set; }

            public DeviceRecordStruct(int id)
            {
                this.Id = id;
                AllInTimes = new int[24];
                AllOutTimes = new int[24];
            }
        }
    }


}