using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using Oracle.ManagedDataAccess.Client;
using Ry.Defined;

namespace 预约签到
{
    public partial class Form1 : Form
    {

        public OracleConnection ORAconn;
        public OracleCommand ORAcmd;
        public OracleDataAdapter ORAadp;
        public string DBIP;
        public string DBName;
        public string DBPasswd;
        public string DBPort;
        public string ServerName;
        public string methodName;
        public string methodCodeType;
        public string methodCodeOther;
        public string methodProgram;
        private string _beep;
        public string configFile = MyConst.ConfigFilePath + "\\Config.xml";
        private DataTable dt;
        private string result="";
        public string ConnectOracle()
        {
            try
            {
                if (ORAconn == null)
                    if (ORAconn == null)
                    {
                        string connString =string.Format(
                            "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={0})" +
                            "(PORT={1}))(CONNECT_DATA=(SERVICE_NAME={2})));" +
                            "Persist Security Info=True;User ID={3};Password={4};",DBIP,DBPort, ServerName, DBName, DBPasswd);
                        ORAconn = new OracleConnection(connString);
                        ORAconn.Open();
                        ORAcmd = new OracleCommand("", ORAconn);
                        ORAadp = new OracleDataAdapter(ORAcmd);
                    }

                return string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return ex.ToString();
            }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Dictionary<string,object> dict = new Dictionary<string, object>();
            dict.Add("method", methodName);
            dict.Add("CodeType", methodCodeType);
            dict.Add("code1", textBox1.Text);
            dict.Add("code2", textBox2.Text);
            dict.Add("other", methodCodeOther);

            ExecuteProgram(methodProgram, dict, out result);
                

            if (_beep != "0")
                Ry.Function.FunctionCommon.MessageBeep(550);
            textBox1.SelectAll();
            textBox1.Focus();

            label4.Text = result;
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            DBIP = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "ConnectDB/DBHostIP", "127.0.0.1");
            DBName = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "ConnectDB/DBUserName", "system"); ;
            DBPasswd = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "ConnectDB/DBPasswd", "Oracle");
            DBPort = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "ConnectDB/DBPort", "1521"); 
            ServerName = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "ConnectDB/DBName", "Oracle"); 
            ConnectOracle();

            dt = Ry.Function.FunctionCommon.XmlFile2Datatable(MyConst.ConfigFilePath + "\\CodeType.xml");

            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "NAME";

            this.Text = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "设置/软件名称", "常用程序设置");
            ConnectOracle();

            methodName = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "设置/处理方式", "执行预约");

            methodCodeOther = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "设置/其他条件", "");

            methodProgram = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "设置/存储过程名", "pro_根据相关的ID执行");

            _beep = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "设置/蜂鸣提示", "1");




            methodCodeType = XMLFileOperate.CXmlFile.GetStringFromXmlFile(configFile,
                "设置/默认操作", "0");

            comboBox1.SelectedValue = methodCodeType;

        }


        public int ExecuteProgram(string programName,Dictionary<string, object> dictParam,out string result)
        {
            result = "";
            ORAcmd.CommandText = programName;
            ORAcmd.CommandType = CommandType.StoredProcedure;
            ORAcmd.Parameters.Clear();
            OracleParameter p1;
            foreach (KeyValuePair<string, object> KeyValue in dictParam)
            {
                p1 = new OracleParameter(KeyValue.Key,OracleDbType.Varchar2);
                p1.Value = KeyValue.Value;
                p1.Direction = ParameterDirection.Input;
                ORAcmd.Parameters.Add(p1);
            }

            OracleParameter p2 = new OracleParameter("ResultStr", OracleDbType.Varchar2,100);
            p2.Direction = ParameterDirection.Output;
            ORAcmd.Parameters.Add(p2);
            try
            {
                ORAcmd.ExecuteNonQuery();
                result = ORAcmd.Parameters["ResultStr"].Value.ToString();
                return MyConst.Success;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return MyConst.Failed;
            }

   
        }

        public  int GetDataSet(string sqlString, CommandType commandType, Dictionary<string, object> dictParam, ref DataTable ds)
        {

            DataSet dt = new DataSet();
            try
            {

                ORAcmd.CommandText = sqlString;
                ORAcmd.CommandType = commandType;

                ORAcmd.Parameters.Clear();
                foreach (KeyValuePair<string, object> KeyValue in dictParam)
                {
                    //OracleParameter p1 = new OracleParameter(KeyValue.Key, OracleType.VarChar);
                    //p1.Value = KeyValue.Value;
                    //p1.Direction = ParameterDirection.Input;
                    //this._command.Parameters.Add(p1);
                    ORAcmd.Parameters.Add(KeyValue.Key, KeyValue.Value);
                }

               // OracleParameter O = new OracleParameter("result", OracleDbType.RefCursor,ParameterDirection.Output);

               // O.Direction = ParameterDirection.Output;
               // ORAcmd.Parameters.Add(O);

                ORAadp.Fill(ds);
            }
            catch (Exception e)
            {
                MyConst.Err.SetError(20002, "GetDataSet - " + sqlString, e.Message, sqlString);
                return MyConst.Failed;
            }

            return MyConst.Success;
        }


        private void Form1_Shown(object sender, EventArgs e)
        {
          
            

        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            methodCodeType = comboBox1.SelectedValue.ToString();
            XMLFileOperate.CXmlFile.SaveStringToXmlFile(comboBox1.SelectedValue.ToString(), configFile, "设置/默认操作");


        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(13))
            {
                button1_Click(sender,e);
            }

        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            textBox1.SelectAll();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Text = comboBox1.Text;
            label1.Text = dt.Rows[comboBox1.SelectedIndex]["Label1"].ToString();
            label2.Text = dt.Rows[comboBox1.SelectedIndex]["Label2"].ToString();
            label3.Text = dt.Rows[comboBox1.SelectedIndex]["Label3"].ToString();
            if (label3.Text == "")
            {
                textBox2.Visible = false;
            }
            else
            {
                textBox2.Visible = true;
            }
            

        }
    }
}
