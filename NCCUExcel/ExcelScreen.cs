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
    public class ExcelScreen
    {
        public static string _TopFolder = "C:\\Users\\jim\\Desktop\\thesis\\";

        public static string _ScreenFileFolder = _TopFolder;

        public static void exportScreen(string fileInputName, string fileOutputName)
        {
            List<ScreenStruct> datas = GetExcel(_ScreenFileFolder + fileInputName);
            List<DeviceRecordStruct> allDeviceTimes = new List<DeviceRecordStruct>();

            for (int i = 0; i < datas.Count; i++)
            {
                if (datas[i].Value.Equals("螢幕打開"))
                    continue;

                if (allDeviceTimes.Find(ById(datas[i].Id)) == null)
                {
                    DeviceRecordStruct deviceRecord = new DeviceRecordStruct(datas[i].Id);
                    allDeviceTimes.Add(deviceRecord);
                }

                DateTime startTime = datas[i - 1].date;
                DateTime endTime = datas[i].date;
                DeviceRecordStruct findDeviceRecord = allDeviceTimes.Find(ById(datas[i].Id));
                FigureOutTimeRange(startTime, endTime, findDeviceRecord.AllTimes);
            }


            using (System.IO.StreamWriter file = new System.IO.StreamWriter(_ScreenFileFolder + fileOutputName))
            {
                file.Write("\t");

                for (int i = 0; i < allDeviceTimes.Count; i++)
                {
                    string ID = allDeviceTimes[i].Id + "\t";
                    file.Write(ID);
                }

                file.WriteLine();

                for (int i = 0; i < 24; i++)
                {
                    string show = i.ToString();

                    for (int j = 0; j < allDeviceTimes.Count(); j++)
                    {
                        show += "\t" + (allDeviceTimes[j].AllTimes[i] / 60);
                    }

                    file.WriteLine(show);
                }
                    
            }
        }


        private static List<ScreenStruct> GetExcel(string path)
        {
            List<ScreenStruct> list = new List<ScreenStruct>();
            using (StreamReader sr = new StreamReader(path))
            {
                String line = sr.ReadToEnd();
                string[] records = line.Split('\n');

                foreach (var record in records)
                {
                    string[] row = record.Split('\t');
                    ScreenStruct one = new ScreenStruct()
                    {
                        Id = Convert.ToInt16(row[0]),
                        Value = row[1].ToString(),
                        date = Convert.ToDateTime(row[2])
                    };

                    list.Add(one);
                }
                
            } 

            return list;
        }

        private static void FigureOutTimeRange(DateTime startTime, DateTime endTime, int[] allTimes)
        {
            if (endTime.Date != startTime.Date)
            { 
                DateTime splitStartTime1 = startTime;
                DateTime splitEndTime1 = new DateTime(startTime.Year, startTime.Month, startTime.Day, 23, 59, 59);
                DateTime splitStartTime2 = new DateTime(endTime.Year, endTime.Month, endTime.Day, 0, 0, 0);
                DateTime splitEndTime2 = endTime;
                FigureOutTimeRange(splitStartTime1, splitEndTime1, allTimes);
                FigureOutTimeRange(splitStartTime2, splitEndTime2, allTimes);
            }
            else if (endTime.Hour != startTime.Hour)
            {
                int interval = endTime.Hour - startTime.Hour;
                for (int i = 0; i <= interval; i++)
                {
                    if (i == 0)
                    {
                        DateTime splitStartTime = startTime;
                        DateTime splitEndTime = new DateTime(endTime.Year, endTime.Month, endTime.Day,
                            startTime.Hour, 59, 59);
                        FigureOutTimeRange(splitStartTime, splitEndTime, allTimes);
                    }
                    else if (i == interval)
                    {
                        DateTime splitStartTime = new DateTime(endTime.Year, endTime.Month, endTime.Day,
                            (startTime.Hour + i), 0, 0);
                        DateTime splitEndTime = endTime;
                        FigureOutTimeRange(splitStartTime, splitEndTime, allTimes);
                    }
                    else
                    {
                        DateTime splitStartTime = new DateTime(endTime.Year, endTime.Month, endTime.Day,
                            (startTime.Hour + i), 0, 0);
                        DateTime splitEndTime = new DateTime(endTime.Year, endTime.Month, endTime.Day,
                            (startTime.Hour + i), 59, 59);
                        FigureOutTimeRange(splitStartTime, splitEndTime, allTimes);
                    }
                }
            }
            else
            {
                int diffSeconds = (endTime - startTime).Minutes * 60 + (endTime - startTime).Seconds;

                allTimes[startTime.Hour] += diffSeconds;
            }
        }

        static Predicate<DeviceRecordStruct> ById(int id)
        {
            return delegate(DeviceRecordStruct deviceRecord)
            {
                return deviceRecord.Id == id;
            };
        }

        private class ScreenStruct
        {
            public int Id { get; set; }
            public string Value { get; set; }
            public DateTime date { get; set; }
        }

        private class DeviceRecordStruct
        {
            public int Id { get; set; }
            public int[] AllTimes { get; set; }

            public DeviceRecordStruct(int id)
            {
                this.Id = id;
                AllTimes = new int[24];
            }
        }
    }


}