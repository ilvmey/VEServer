using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using System.Data.SQLite;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.Office.Interop.Excel;
using myexcelcollection;
using veclient;
using vesocketserver;
using downloadpath;
using OnlineClient;
namespace VEServer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        _Application myExcel=null;
        _Workbook myBook=null;
        _Worksheet mySheet=null;
        Range myRange=null;
        DownloadPath dlPath = new DownloadPath();
        MyExcelCollection[] question=null;
        Socket[] sckAccept;
        VESocket[] SckSs;
        string LocalIP;
        int port = 1234;
        int SckCIndex = 0;
        VEClient vec=new VEClient();
        OpenExam openExam = new OpenExam();

        int accumulate = 0;//累積連線人數
        OC onlineClient = new OC();//在線人數物件
        int online = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            lblHostName.Text = Dns.GetHostName();
            IPAddress[] ipa = Array.FindAll(Dns.GetHostEntry(string.Empty).AddressList, a => a.AddressFamily == AddressFamily.InterNetwork);
            LocalIP = ipa[0].ToString();
            lblHostIP.Text = "本機IP=" + ipa[0].ToString();
            Array.Resize(ref SckSs, 1);
            Array.Resize(ref sckAccept, 1);
            sckAccept[0] = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            sckAccept[0].Bind(new IPEndPoint(IPAddress.Parse(LocalIP), port));
            sckAccept[0].Listen(5);
            SckSs[0] = new VESocket();
            //SckSs[0].setSck();
            SckSsWaitAccept();
        }

        private void SckSsWaitAccept()
        {
            SckCIndex = SckSs.Length;
            Array.Resize(ref SckSs, SckCIndex + 1);
            Array.Resize(ref sckAccept, SckCIndex + 1);
            Thread SckSAcceptTd = new Thread(SckSAcceptProc);
            SckSAcceptTd.Start();
        }

        private void SckSAcceptProc()
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            int Scki = SckCIndex;
            SckSs[Scki] = new VESocket();
            sckAccept[Scki] = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            try
            {
                //Socket s = SckSs[Scki].getSck();
                //s = sckAccept.Accept();
                //SckSs[Scki].SckAccept(sckAccept);
                sckAccept[Scki] = sckAccept[0].Accept();
                accumulate++;
                lblAcc.Text = Convert.ToString(accumulate);
                //SckSs[Scki].setSck(s);
                SckSsWaitAccept();
                byte[] clientData = new byte[1024];
                while (true)
                {
                   // if (SckSs[Scki].getSck().Connected == true)
                    if (sckAccept[Scki].Connected == true)
                    {                      
                        //SckSs[Scki].getSck().Receive(clientData);
                        sckAccept[Scki].Receive(clientData);
                        BinaryFormatter bf = new BinaryFormatter();
                        MemoryStream stream = new MemoryStream(clientData);

                        object obj = bf.Deserialize(stream);          
                        if (obj.GetType() == vec.GetType())
                        {
                            vec = (VEClient)obj;
                            //lblclientComputerName.Text = vec.getComputerName();
                            lblClientIpAddress.Text = vec.getIpAddress();
                            listBox1.Items.Add(vec.getComputerName());
                            //SckSs[0];
                        }
                        else if(obj.GetType()==dlPath.GetType())
                        {
                            dlPath = (DownloadPath)obj;
                            sendMyExcelCollection(dlPath.getPath(),Scki);
                        }

                    }
                }
            }
            catch { }
        }
        void sendMyExcelCollection(string path,int scki)
        {
            question = openExam.open(path);
            Serialize(question, scki);
        }


        void openExcel(string path)
        {
            myExcel = new Microsoft.Office.Interop.Excel.Application();
            myExcel.Workbooks.Open(path);
            myExcel.DisplayAlerts = false;
            myExcel.Visible = false;
            myBook = myExcel.Workbooks[1];
            myBook.Activate();
            mySheet = (_Worksheet)myBook.Worksheets[1];
            mySheet.Activate();

        }
        void readExcel()
        {

            int count = 1;
            string raws = "A" + count;
            myRange = mySheet.get_Range(raws);
            Array.Resize(ref question, 1);
            while (Convert.ToString(myRange.Value) != null)
            {
                raws = "A" + count;
                myRange = mySheet.get_Range(raws);
                string q = Convert.ToString(myRange.Value);

                raws = "B" + count;
                myRange = mySheet.get_Range(raws);
                string a = Convert.ToString(myRange.Value);

                raws = "C" + count;
                myRange = mySheet.get_Range(raws);
                string b = Convert.ToString(myRange.Value);

                raws = "D" + count;
                myRange = mySheet.get_Range(raws);
                string c = Convert.ToString(myRange.Value);
                raws = "E" + count;
                myRange = mySheet.get_Range(raws);
                string d = Convert.ToString(myRange.Value);

                raws = "F" + count;
                myRange = mySheet.get_Range(raws);
                string ans = Convert.ToString(myRange.Value);
                question[count - 1] = new MyExcelCollection(q, a, b, c, d, ans);
                Array.Resize(ref question, question.Length + 1);

                raws = "A" + ++count;
                myRange = mySheet.get_Range(raws);
            }
            Array.Resize(ref question, question.Length - 1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
            myBook = null;
            mySheet = null;
            myRange = null;
            myExcel = null;
            GC.Collect();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach(VESocket s in SckSs)
            {
                s.getSck().Close();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            online = 0;
            onlineClient.SetOnlineClient(0);
            listBox2.Items.Clear();
            for(int i=1;i<SckSs.Length-1;i++)
            {
                string s = "SocketIndex:" + i + " Connected:" + sckAccept[i].Connected;
                listBox2.Items.Add(s);
                if (sckAccept[i].Connected)
                    online++;
            }
            lblSckIndex.Text = Convert.ToString(SckCIndex);
           
            lblOnline.Text = Convert.ToString(online);
            onlineClient.SetOnlineClient(online);
            OnlineClientBroadcast();
        }

        public void Serialize(MyExcelCollection[] m,int scki)
        {
            BinaryFormatter bf = new BinaryFormatter();
            MemoryStream stream = new MemoryStream();
            bf.Serialize(stream, m);
            byte[] bytesSend = new byte[1024];
            bytesSend = stream.ToArray();
            sckAccept[scki].Send(bytesSend);
        }
        public void OnlineClientBroadcast()
        {
            for(int i=1;i<sckAccept.Length-1;i++)
            {
                if (sckAccept[i].Connected)
                    Serialize(onlineClient,i);
            }
        }
        public void Serialize(OC oc,int scki)
        {
            /*BinaryFormatter bf = new BinaryFormatter();
            MemoryStream stream = new MemoryStream();
            bf.Serialize(stream, oc);
            byte[] bytesSend = new byte[1024];
            bytesSend = stream.ToArray();
            sckAccept[scki].Send(bytesSend);*/
        }
    }
}
