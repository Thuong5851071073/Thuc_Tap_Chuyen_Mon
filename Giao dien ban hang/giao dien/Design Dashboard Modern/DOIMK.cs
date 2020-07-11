    using Design_Dashboard_Modern.DAO;
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

namespace Design_Dashboard_Modern
{
    public partial class DOIMK : Form
    {
        public DOIMK()
        {
            InitializeComponent();
        }

        #region METHOD:

        #region METHOD-CHECK-OLD-PASSWORD:

        bool checkoldpass(string oldpass,string gmail)
        {
           return AccountDAO.Instance.DAO_CHECKMK(gmail,oldpass);
        }

        #endregion

        #region METHOD-CHANGE-PASSWORD:

        void Update(string newpassword, string gmail, string oldpass)
        {
            byte[] temp = ASCIIEncoding.ASCII.GetBytes(newpassword);
            byte[] hasdata = new MD5CryptoServiceProvider().ComputeHash(temp);
            string haspass = "";
            foreach (byte item in hasdata)
            {
                haspass += item;
            }
            if (checkoldpass(gmail, oldpass) == true)
            {
                {
                    MessageBox.Show("Cập Nhật tài khoản thành công");
                }
            }
            else
            {
                MessageBox.Show("Mật khẩu cũ không đúng, Vui lòng nhập lại!");

            }
            /*if (checkoldpass(gmail,oldpass ) == AccountDAO.Instance.DAO_updateAccount(haspass, gmail))
            {
                MessageBox.Show("UpDate tài khoản thành công");
            }
            
            else
            {
                MessageBox.Show("Mật khẩu cũ không đúng, Vui lòng nhập lại!");

            }*/

        }
        #endregion

        #endregion

        #region EVENT:

        #region EVENT-ADD-ACCOUNT:

       
        private void btnChange_Click_1(object sender, EventArgs e)
        {
            string newpassword1 = tbxNewPass1.Text;
          
            string gmail = tbxGmail.Text;
            string oldpass = tbxOldPassWord.Text;
            if (oldpass != newpassword1 )
            {
              
                    Update(newpassword1, gmail, oldpass);
                       
              
            }
            else
            {
                MessageBox.Show("Vui lòng nhập mật khẩu mới khác mật khẩu cũ");
            }
        }
        #endregion

        #region EVENT-CLOSE-FROM:
        private void btnCancle_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }
       
        private void DOIMK_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Bạn có thật sự muốn rời khỏi đây", "Thông báo", MessageBoxButtons.OKCancel) != System.Windows.Forms.DialogResult.OK)
            {
                e.Cancel = true;
            }
        }

        #endregion

        #region EVENT-HIDE-TEXT:

      
        private void tbxGmail_Enter_1(object sender, EventArgs e)
        {
          
            {
                tbxGmail.Text = "";
            }
        }

        private void tbxGmail_Leave_1(object sender, EventArgs e)
        {
            if (tbxGmail.Text == "")
            {
                tbxGmail.Text = "TÊN ĐĂNG NHẬP";
            }
        }

        private void tbxOldPassWord_Enter_1(object sender, EventArgs e)
        {
         
            {
                tbxOldPassWord.Text = "";
            }
            
        }
        private void tbxOldPassWord_Leave_1(object sender, EventArgs e)
        {
            if (tbxOldPassWord.Text == "")
            {
                
                tbxOldPassWord.Text = "MẬT KHẨU CŨ";
            }
        }
        private void tbxNewPass1_Enter_1(object sender, EventArgs e)
        {
      
            {
                tbxNewPass1.Text = "";
            }
            

        }

        private void tbxNewPass1_Leave_1(object sender, EventArgs e)
        {
            if (tbxNewPass1.Text == "")
            {
                
                tbxNewPass1.Text = "MẬT KHẨU MỚI";
            }
        }
       

      

 
        #endregion

        #endregion

        private void DOIMK_Load(object sender, EventArgs e)
        {
            panel1.BackColor = Color.FromArgb(100, 0, 0, 0);
           
        }

       
    }
}
