using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace report2
{
    public partial class Login : Form
    {
        int times;
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            
        }
        

       
        //登入查詢
        private void LoginLoad_Click(object sender, EventArgs e)
        {

            string name = "", status = "", hashpwd = "",catsys = "";
            hashpwd = getMd5Hash(password.Text);
            MySqlConnection cnn = new MySqlConnection("server=125.227.18.121;database=eversta2;uid=reportuser;pwd=2uzixydu@;");
            MySqlCommand cmd = new MySqlCommand("SELECT u.`name`,u.`status`,u.catsys from `user` u where  account = '" + account.Text + "' and password = '" + hashpwd + "';", cnn);
            cnn.Open();
            MySqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                name = (string)dr.GetValue(0);
                status = (string)dr.GetValue(1);
                catsys = (string)dr.GetValue(2);
            }
            if(status != ""  )
            {
                if(status !="0" && catsys != "0")
                {
                    this.Hide();
                    //傳值給第二視窗
                    Report_health_check rhc = new Report_health_check();
                    rhc.UserName = name;
                    rhc.setUserValue();
                    rhc.ShowDialog(this);

                    password.Text = "";
                }
                else
                {
                    MessageBox.Show("您的帳號已失效，請聯絡資訊人員");
                }

                
            }else
            {
                times++;
                if (times == 3)
                {
                    MessageBox.Show("登入失敗三次\n請確認帳號密碼");
                    Environment.Exit(Environment.ExitCode);
                }
                MessageBox.Show("登入失敗");
            }
            cnn.Close();
        }
        //md5加密
        protected String getMd5Hash(String Input_text)
        {
            MD5 md5Hasher = MD5.Create();
            Byte[] MD5data = md5Hasher.ComputeHash(Encoding.Default.GetBytes(Input_text));
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < (MD5data.Length); i++)
            {
                sBuilder.Append(MD5data[i].ToString("x2"));  //--變成十六進位
            }
            return sBuilder.ToString();
        }

        private void password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoginLoad_Click(this, new EventArgs());
            }
        }
    }
}
