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
            statusStrip1.Items[2].Text = MyModule.username;
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void UNB_NodeMouseClick(TreeNode node, int menuIndex, int pageIndex)
        {
            
        }
    }
}
