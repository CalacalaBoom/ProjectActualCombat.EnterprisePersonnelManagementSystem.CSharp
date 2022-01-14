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
using Sunny.UI;
using EPMS.ModuleClass;
using System.IO;

namespace EPMS
{
    public partial class F_FileMan : UIForm
    {
        public void saveTo(tb_staffbasic _Staffbasic)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                string index = db.tb_staffbasic.Max(M => M.ID);
                if (index != null)
                {
                    txtID.Text = (Convert.ToInt32(index) + 1).ToString();
                }
                else
                {
                    txtID.Text = "1";
                }
                _Staffbasic.ID = txtID.Text;
                _Staffbasic.StaffName = txtName.Text;
                _Staffbasic.Folk = comFolk.Text;
                _Staffbasic.Birthday = uiDatePicker1.Value;
                _Staffbasic.Age = Convert.ToInt32(txtAge.Text);
                _Staffbasic.Culture = comCul.Text;
                _Staffbasic.Marriage = comMarr.Text;
                _Staffbasic.Sex = comSex.Text;
                _Staffbasic.Visage = comVisage.Text;
                _Staffbasic.IDCard = txtIDCard.Text;
                _Staffbasic.Workdate = uiDatePicker2.Value;
                _Staffbasic.WorkLength = Convert.ToInt32(txtWorkLength.Text);
                _Staffbasic.Employee = comEmployee.Text;
                _Staffbasic.Business = comBusiness.Text;
                _Staffbasic.Laborage = comLaborage.Text;
                _Staffbasic.Branch = comBranch.Text;
                _Staffbasic.Duthcall = comDuthCall.Text;
                _Staffbasic.Phone = txtPhone.Text;
                _Staffbasic.Handset = txtHandset.Text;
                _Staffbasic.School = txtSchool.Text;
                _Staffbasic.Speciality = txtSpeciality.Text;
                _Staffbasic.GraduateDate = uiDatePicker5.Value;
                _Staffbasic.Address = txtAddress.Text;
                _Staffbasic.Photo = MyModule.imgBytesln;
                _Staffbasic.BeAware = txtBeAware.Text;
                _Staffbasic.City = txtCity.Text;
                _Staffbasic.M_Pay = Convert.ToSingle(txtPay.Text);
                _Staffbasic.Bank = txtBank.Text;
                _Staffbasic.Pact_B = uiDatePicker3.Value;
                _Staffbasic.Pact_E = uiDatePicker4.Value;
                _Staffbasic.Pact_Y = Convert.ToSingle(txtY.Text);

                //添加SQL
                db.tb_staffbasic.Add(_Staffbasic);
                db.SaveChanges();//进行数据库添加操作

                UIMessageTip.ShowOk("保存成功");
            }
        }

        public void readOut(tb_staffbasic _Staffbasic)
        {
            txtID.Text = _Staffbasic.ID;
            txtName.Text = _Staffbasic.StaffName;
            comFolk.Text = _Staffbasic.Folk;
            uiDatePicker1.Value = _Staffbasic.Birthday.Value;
            txtAge.Text = _Staffbasic.Age.ToString();
            comCul.Text = _Staffbasic.Culture;
            comMarr.Text = _Staffbasic.Marriage;
            comSex.Text = _Staffbasic.Sex;
            comVisage.Text = _Staffbasic.Visage;
            txtIDCard.Text = _Staffbasic.IDCard;
            uiDatePicker2.Value = _Staffbasic.Workdate.Value;
            txtWorkLength.Text = _Staffbasic.WorkLength.ToString();
            comEmployee.Text = _Staffbasic.Employee;
            comBusiness.Text = _Staffbasic.Business;
            comLaborage.Text = _Staffbasic.Laborage;
            comBranch.Text = _Staffbasic.Branch;
            comDuthCall.Text = _Staffbasic.Duthcall;
            txtPhone.Text = _Staffbasic.Phone;
            txtHandset.Text = _Staffbasic.Handset;
            txtSchool.Text = _Staffbasic.School;
            txtSpeciality.Text = _Staffbasic.Speciality;
            uiDatePicker5.Value = _Staffbasic.GraduateDate.Value;
            txtAddress.Text = _Staffbasic.Address;
            MyModule.imgBytesln = _Staffbasic.Photo;
            txtBeAware.Text = _Staffbasic.BeAware;
            txtCity.Text = _Staffbasic.City;
            txtPay.Text = _Staffbasic.M_Pay.ToString();
            txtBank.Text = _Staffbasic.Bank;
            uiDatePicker3.Value = _Staffbasic.Pact_B.Value;
            uiDatePicker4.Value = _Staffbasic.Pact_E.Value;
            txtY.Text = _Staffbasic.Pact_Y.ToString();
            if (MyModule.imgBytesln != null)
            {
                MemoryStream ms = new MemoryStream(MyModule.imgBytesln);
                Bitmap bmpt = new Bitmap(ms);
                picPhoto.Image = bmpt;
            }
            else
            {
                picPhoto.Image = null;
            }
        }

        public void change(String index)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                tb_staffbasic _Staffbasic = new tb_staffbasic { ID = index };
                db.tb_staffbasic.Attach(_Staffbasic);

                _Staffbasic.StaffName = txtName.Text;
                _Staffbasic.Folk = comFolk.Text;
                _Staffbasic.Birthday = uiDatePicker1.Value;
                _Staffbasic.Age = Convert.ToInt32(txtAge.Text);
                _Staffbasic.Culture = comCul.Text;
                _Staffbasic.Marriage = comMarr.Text;
                _Staffbasic.Sex = comSex.Text;
                _Staffbasic.Visage = comVisage.Text;
                _Staffbasic.IDCard = txtIDCard.Text;
                _Staffbasic.Workdate = uiDatePicker2.Value;
                _Staffbasic.WorkLength = Convert.ToInt32(txtWorkLength.Text);
                _Staffbasic.Employee = comEmployee.Text;
                _Staffbasic.Business = comBusiness.Text;
                _Staffbasic.Laborage = comLaborage.Text;
                _Staffbasic.Branch = comBranch.Text;
                _Staffbasic.Duthcall = comDuthCall.Text;
                _Staffbasic.Phone = txtPhone.Text;
                _Staffbasic.Handset = txtHandset.Text;
                _Staffbasic.School = txtSchool.Text;
                _Staffbasic.Speciality = txtSpeciality.Text;
                _Staffbasic.GraduateDate = uiDatePicker5.Value;
                _Staffbasic.Address = txtAddress.Text;
                _Staffbasic.Photo = MyModule.imgBytesln;
                _Staffbasic.BeAware = txtBeAware.Text;
                _Staffbasic.City = txtCity.Text;
                _Staffbasic.M_Pay = Convert.ToSingle(txtPay.Text);
                _Staffbasic.Bank = txtBank.Text;
                _Staffbasic.Pact_B = uiDatePicker3.Value;
                _Staffbasic.Pact_E = uiDatePicker4.Value;
                _Staffbasic.Pact_Y = Convert.ToSingle(txtY.Text);

                db.SaveChanges();

                UIMessageTip.ShowOk("修改成功");
            }
        }

        public void delete(String index)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index).FirstOrDefault();
                db.tb_staffbasic.Remove(_Staffbasic);
                db.SaveChanges();
                UIMessageTip.ShowOk("删除成功");
            }
        }

        public F_FileMan()
        {
            InitializeComponent();
        }

        private void uiDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_showall_Click(object sender, EventArgs e)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                uiDataGridView1.DataSource = db.tb_staffbasic.Select(S => new { S.ID, S.StaffName }).ToList();
            }
        }

        private void F_FileMan_Load(object sender, EventArgs e)
        {

        }

        private void uiDatePicker1_TextChanged(object sender, EventArgs e)
        {
            txtAge.Text = (DateTime.Today.Year - uiDatePicker1.Year + 1).ToString();
        }

        private void uiDatePicker2_ValueChanged(object sender, DateTime value)
        {
            txtWorkLength.Text = (DateTime.Today.Year - uiDatePicker2.Year).ToString();
        }

        private void uiDatePicker4_ValueChanged(object sender, DateTime value)
        {
            if (uiDatePicker4.Text != null && uiDatePicker3.Text != null)
            {
                txtY.Text = (uiDatePicker4.Year - uiDatePicker3.Year).ToString();
            }
        }

        private void uiDatePicker3_ValueChanged(object sender, DateTime value)
        {
            if (uiDatePicker4.Text != null && uiDatePicker3.Text != null)
            {
                txtY.Text = (uiDatePicker4.Year - uiDatePicker3.Year).ToString();
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog photoFileDialog = new OpenFileDialog();
            photoFileDialog.InitialDirectory = "C:\\";//初始目录为C盘根目录
            MyModule imgModule = new MyModule();
            imgModule.Read_Image(photoFileDialog, picPhoto, txtphotoPath);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            picPhoto.Image = null;
            txtphotoPath.Text = String.Empty;
        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                tb_staffbasic _Staffbasic = new tb_staffbasic();
                saveTo(_Staffbasic);
            }
        }

        private void uiDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = (string)uiDataGridView1.CurrentRow.Cells[1].Value;
            int index = uiDataGridView1.CurrentRow.Index + 1;
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index.ToString()).FirstOrDefault();
                readOut(_Staffbasic);


            }
        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            if (txtID.Text == string.Empty)
            {
                UIMessageTip.ShowError("请先在左侧选择要修改的档案");
            }
            else
            {
                String index = txtID.Text;
                change(index);
                using (db_pwmsEntities db = new db_pwmsEntities())
                {
                    uiDataGridView1.DataSource = db.tb_staffbasic.Select(S => new { S.ID, S.StaffName }).ToList();
                }
            }
        }

        private void btn_Delete_Click(object sender, EventArgs e)
        {
            if (txtID.Text == string.Empty)
            {
                UIMessageTip.ShowError("请先在左侧选择要删除的档案");
            }
            else
            {
                String index = txtID.Text;
                delete(index);
                using (db_pwmsEntities db = new db_pwmsEntities())
                {
                    uiDataGridView1.DataSource = db.tb_staffbasic.Select(S => new { S.ID, S.StaffName }).ToList();
                }
            }
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void uiComboBox1_TextChanged(object sender, EventArgs e)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                switch (uiComboBox1.SelectedIndex)
                {
                    case 0:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.StaffName).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 1:
                        {
                            uiComboBox2.Items.Clear();
                            uiComboBox2.Items.Add("男");
                            uiComboBox2.Items.Add("女");
                            break;
                        }
                    case 2:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Folk).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 3:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Culture).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 4:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Visage).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 5:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Employee).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 6:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Business).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 7:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Branch).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 8:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Duthcall).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                    case 9:
                        {
                            uiComboBox2.Items.Clear();
                            foreach (string i in db.tb_staffbasic.Select(S => S.Laborage).Distinct().ToList())
                            {
                                uiComboBox2.Items.Add(i);
                            }
                            break;
                        }
                }
            }

        }

        private void uiComboBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string selected = uiComboBox2.SelectedItem.ToString();
                using (db_pwmsEntities db = new db_pwmsEntities())
                {
                    if (uiComboBox1.SelectedIndex == 0)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.StaffName == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 1)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Sex == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 2)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Folk == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 3)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Culture == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 4)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Visage == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 5)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Employee == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 6)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Business == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 7)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Branch == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else if (uiComboBox1.SelectedIndex == 8)
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Duthcall == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                    else
                        uiDataGridView1.DataSource = db.tb_staffbasic.Where(W => W.Laborage == uiComboBox2.Text).Select(S => new { S.ID, S.StaffName }).ToList();
                }
            }
            catch
            {
                uiComboBox2.Text = "";
                UINotifier.Show("只能以选择的方式查询", UINotifierType.ERROR);
            }
        }

        private void N_First_Click(object sender, EventArgs e)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == "1").FirstOrDefault();
                if (_Staffbasic != null)
                    readOut(_Staffbasic);
            }
        }

        private void N_Cauda_Click(object sender, EventArgs e)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                string index = db.tb_staffbasic.Max(M => M.ID);
                tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index).FirstOrDefault();
                if (_Staffbasic != null)
                    readOut(_Staffbasic);
            }

        }

        private void N_Pre_Click(object sender, EventArgs e)
        {
            if (txtID.Text == "")
            {
                UINotifier.Show("没有正在浏览的档案", UINotifierType.ERROR);
            }
            else if (txtID.Text == "1")
            {
                UINotifier.Show("已经是第一张档案了", UINotifierType.ERROR);
            }
            else
            {
                string index = (Convert.ToInt32(txtID.Text) - 1).ToString();
                using (db_pwmsEntities db = new db_pwmsEntities())
                {
                    tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index).FirstOrDefault();
                    if (_Staffbasic != null)
                        readOut(_Staffbasic);
                }
            }
        }

        private void N_Next_Click(object sender, EventArgs e)
        {
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                string last = db.tb_staffbasic.Max(M => M.ID);
                if (txtID.Text == "")
                {
                    UINotifier.Show("没有正在浏览的档案", UINotifierType.ERROR);
                }
                else if (txtID.Text == last)
                {
                    UINotifier.Show("已经是最后一张档案了", UINotifierType.ERROR);
                }
                else
                {
                    string index = (Convert.ToInt32(txtID.Text) + 1).ToString();
                    tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index).FirstOrDefault();
                    if (_Staffbasic != null)
                        readOut(_Staffbasic);
                }
            }
        }

        private void btn_toWord_Click(object sender, EventArgs e)
        {
            object Nothing = System.Reflection.Missing.Value;
            object missing = System.Reflection.Missing.Value;
            //创建文档
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            wordApp.Visible = true;
            //设置文档宽度
            wordApp.Selection.PageSetup.LeftMargin = wordApp.CentimetersToPoints(float.Parse("2"));
            wordApp.ActiveWindow.ActivePane.HorizontalPercentScrolled = 11;
            wordApp.Selection.PageSetup.RightMargin = wordApp.CentimetersToPoints(float.Parse("2"));
            Object start = Type.Missing;
            Object end = Type.Missing;
            PictureBox pp = new PictureBox();
            /*int p1 = 0;
            for (int i = 0; i < ; i++)
            {

            }*/
            object rng = Type.Missing;
            string strinfo = "职工基本信息表" + "(" + txtName.Text + ")";
            start = 0;
            end = 0;
            wordDoc.Range(ref start, ref end).InsertBefore(strinfo);
            wordDoc.Range(ref start, ref end).Font.Name = "Verdana";
            wordDoc.Range(ref start, ref end).Font.Size = 20;



        }
    }
}

