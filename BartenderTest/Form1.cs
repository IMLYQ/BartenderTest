using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BartenderTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string serialNumber = textBox1.Text;
            string origin = textBox2.Text;
            StringBuilder sb = new StringBuilder();
            Type t = typeof(Info);
            PropertyInfo[] infos = t.GetProperties();
            string att = string.Empty;
            for (int i = 0; i < infos.Length; i++)
            {
                if (i == infos.Length - 1)
                {
                    att += "\"" + infos[i].Name + "\"";
                }
                else
                {
                    att += "\"" + infos[i].Name + "\",";
                }
            }
            sb.AppendLine(att);

            //拼接数据
            sb.AppendLine(serialNumber + "," + origin);

            SaveToTxt(sb, path + "\\$BarCodeTest003.txt"); 
            string filePath = path + "\\BarCodeTest003";
            PrintMethod(filePath, "1");
        }


        /// <summary>
        /// BarTender列印方法1
        /// </summary>
        /// <param name="destFilePath">文件路径</param>
        /// <param name="copies">列印份数</param>
        public void PrintMethod(string destFilePath, string copies)
        {

            BarTender.Application btApp = new BarTender.Application();
            /// <summary>
            /// 打印管理对象
            /// </summary>
            BarTender.Format btFormat;
            btFormat = btApp.Formats.Open(destFilePath, false, "");
            btFormat.InputDataFileSetup.TextFile.FileName = (AppDomain.CurrentDomain.BaseDirectory + "$BarCodeTest003.txt");
            btFormat.PrintSetup.UseDatabase = true;
            btFormat.PrintSetup.IdenticalCopiesOfLabel = 1;  //设置同序列打印的份数
            btFormat.PrintSetup.NumberSerializedLabels = 1;  //设置需要打印的序列数 

            if (btFormat.PrintSetup.UseDatabase)
            {
                btFormat.RecordRange = "1...";
            }
            //BarTenderSetParms(btFormat, Parms);
            int A = btFormat.PrintOut(false, true); //第二个false设置打印时是否跳出打印属性
            btFormat.Close(BarTender.BtSaveOptions.btDoNotSaveChanges); //退出时是否保存标签



            Process p = new Process();
            p.StartInfo.FileName = "bartend.exe";
            //列印btw檔案並最小化程序
            p.StartInfo.Arguments = $@"/AF={destFilePath} /P /min=SystemTray";
            p.EnableRaisingEvents = true;
            int pageCount = Convert.ToInt32(copies);
            for (int i = 0; i < pageCount; i++)
            {
                p.Start();
            }
        }


        /// <summary>
        /// 保存txt文本
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private void SaveToTxt(StringBuilder sb, string filePath)
        {
            try
            {
                FileStream fileStream = new FileStream(filePath, FileMode.Create);
                //使用UTF8格式防止乱码
                StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8);
                streamWriter.Write(sb);
                streamWriter.Flush();
                streamWriter.Close();
                fileStream.Close(); 
            }
            catch (Exception e)
            {
            }
        }
        /// <summary>
        /// 打印信息
        /// </summary>
        public class Info
        {
            public string Number { get; set; }

            public string Address { get; set; }
        }
    }
}
