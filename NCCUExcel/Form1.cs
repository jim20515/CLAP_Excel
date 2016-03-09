using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NCCUExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //ExcelScreen.exportScreen("screen.csv", "screenOut.csv");
            ExcelCall.exportCall("call.csv", "callIn.csv", "callOut.csv");
            //ExcelApp.exportApp("app.csv", "appOut.csv");

            //for (int i = 0; i < _CallFileList.Count(); i++)
            //{
            //    ExcelCall.exportCall(_CallFileList[i], _CallFileOutList[i]);
            //}

            //for (int i = 0; i < _AppFileList.Count(); i++)
            //{
            //    ExcelApp.exportApp(_AppFileList[i], _AppFileOutList[i]);
            //}
        }
    }
}
