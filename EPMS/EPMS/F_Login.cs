using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPMS.DataEntity;
using EPMS.ModuleClass;
using Sunny.UI;

namespace EPMS
{
    public partial class F_Login : Form
    {
        public F_Login()
        {
            InitializeComponent();
        }

        

        private void btnLogin_Click(object sender, EventArgs e)
        {
                using (db_pwmsEntities db = new db_pwmsEntities())
                {
                    tb_login _loginUser = db.tb_login.Where(W => W.Name == txtName.Text).FirstOrDefault();
                    if (_loginUser != null)
                    {
                        tb_login _loginPass = db.tb_login.Where(W => W.Pass == txtPass.Text).FirstOrDefault();
                        if (_loginPass != null)
                        {
                            MyModule.FormHelper = true;
                            MyModule.username = txtName.Text;
                            Close();
                        }
                        else
                        {
                            UIMessageTip.Show("密码错误", TipStyle.Red);
                            txtName.Text = null;
                            txtPass.Text = null;
                        }
                    }
                    else
                    {
                        UIMessageTip.Show("没有该用户", TipStyle.Red);
                        txtName.Text = null;
                        txtPass.Text = null;
                    }
                }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void txtName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
                txtPass.Focus();
        }

        private void txtPass_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
                btnLogin.Focus();
        }
    }
}
