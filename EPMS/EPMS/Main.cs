using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Sunny.UI;
using EPMS.ModuleClass;

namespace EPMS
{
    public partial class Main : UIForm
    {
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void UNB_NodeMouseClick(TreeNode node, int menuIndex, int pageIndex)
        {
            
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MyModule.mainFormCon = false;
            new F_FileMan().Show();
        }

        private void UNM_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            

        }

        private void UNM_MenuItemClick(TreeNode node, NavMenuItem item, int pageIndex)
        {
            UIMessageTip.ShowOk(node.Text);
            if (node.Name=="Node37")
            {
                MyModule.mainFormCon = false;
                new F_FileMan().Show();

            }
        }

        private void UNB_MenuItemClick(string itemText, int menuIndex, int pageIndex)
        {
            UIMessageTip.ShowOk(itemText);
            if (itemText == "人事档案管理")
            {
                MyModule.mainFormCon = false;
                new F_FileMan().Show();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Tick += delegate
            {
                if (MyModule.mainFormCon == false)
                {
                    this.WindowState = (FormWindowState)1;
                }
                else
                {
                    this.WindowState = (FormWindowState)2;
                }
            };
            timer1.Start();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.MdiParent = this;
            
            form1.Show();
        }
    }
}
