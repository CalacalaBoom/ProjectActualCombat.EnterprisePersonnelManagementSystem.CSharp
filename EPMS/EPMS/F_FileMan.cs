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
using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using MSExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Threading;

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

        ~F_FileMan()
        {
            MyModule.mainFormCon = true;
        }

        private void uiDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_showall_Click(object sender, EventArgs e)
        {
            UIWaitFormService.ShowWaitForm("正在载入档案列表请稍后...");
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                uiDataGridView1.DataSource = db.tb_staffbasic.Select(S => new { S.ID, S.StaffName }).ToList();
                uiDataGridView1.Columns[0].HeaderText = "职工编号";
                uiDataGridView1.Columns[1].HeaderText = "职工姓名";
                UIWaitFormService.HideWaitForm();
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
            UIWaitFormService.ShowWaitForm("正在载入请稍后...");
            textBox1.Text = (string)uiDataGridView1.CurrentRow.Cells[1].Value;
            int index = uiDataGridView1.CurrentRow.Index + 1;
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index.ToString()).FirstOrDefault();
                readOut(_Staffbasic);
                UIWaitFormService.HideWaitForm();

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
            MyModule.mainFormCon = true;
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
            UIWaitFormService.ShowWaitForm("正在查询，请稍后...");
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
                    UIWaitFormService.HideWaitForm();
                }
            }
            catch
            {
                uiComboBox2.Text = "";
                UINotifier.Show("只能以选择的方式查询", UINotifierType.ERROR);
                UIWaitFormService.HideWaitForm();
            }
        }

        private void N_First_Click(object sender, EventArgs e)
        {
            UIWaitFormService.ShowWaitForm("系统正在处理中...");
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == "1").FirstOrDefault();
                if (_Staffbasic != null)
                {
                    readOut(_Staffbasic);
                    UIWaitFormService.HideWaitForm();
                }
                else
                {
                    UIWaitFormService.HideWaitForm();
                    UIMessageBox.ShowError("找不到档案");
                }
                    
            }
        }

        private void N_Cauda_Click(object sender, EventArgs e)
        {
            UIWaitFormService.ShowWaitForm("系统正在处理中...");
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                string index = db.tb_staffbasic.Max(M => M.ID);
                tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index).FirstOrDefault();
                if(_Staffbasic != null)
                {
                    readOut(_Staffbasic);
                    UIWaitFormService.HideWaitForm();
                }
                else
                {
                    UIWaitFormService.HideWaitForm();
                    UIMessageBox.ShowError("找不到档案");
                }
            }

        }

        private void N_Pre_Click(object sender, EventArgs e)
        {
            UIWaitFormService.ShowWaitForm("系统正在处理中...");
            if (txtID.Text == "")
            {
                UINotifier.Show("没有正在浏览的档案", UINotifierType.ERROR);
                UIWaitFormService.HideWaitForm();
            }
            else if (txtID.Text == "1")
            {
                UINotifier.Show("已经是第一张档案了", UINotifierType.ERROR);
                UIWaitFormService.HideWaitForm();
            }
            else
            {
                string index = (Convert.ToInt32(txtID.Text) - 1).ToString();
                using (db_pwmsEntities db = new db_pwmsEntities())
                {
                    tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index).FirstOrDefault();
                    if (_Staffbasic != null)
                    {
                        readOut(_Staffbasic);
                        UIWaitFormService.HideWaitForm();
                    }
                    else
                    {
                        UIWaitFormService.HideWaitForm();
                        UIMessageBox.ShowError("找不到档案");
                    }
                }
            }
        }

        private void N_Next_Click(object sender, EventArgs e)
        {
            UIWaitFormService.ShowWaitForm("系统正在处理中...");
            using (db_pwmsEntities db = new db_pwmsEntities())
            {
                string last = db.tb_staffbasic.Max(M => M.ID);
                if (txtID.Text == "")
                {
                    UINotifier.Show("没有正在浏览的档案", UINotifierType.ERROR);
                    UIWaitFormService.HideWaitForm();
                }
                else if (txtID.Text == last)
                {
                    UINotifier.Show("已经是最后一张档案了", UINotifierType.ERROR);
                    UIWaitFormService.HideWaitForm();
                }
                else
                {
                    string index = (Convert.ToInt32(txtID.Text) + 1).ToString();
                    tb_staffbasic _Staffbasic = db.tb_staffbasic.Where(W => W.ID == index).FirstOrDefault();
                    if (_Staffbasic != null)
                    {
                        readOut(_Staffbasic);
                        UIWaitFormService.HideWaitForm();
                    }
                    else
                    {
                        UIWaitFormService.HideWaitForm();
                        UIMessageBox.ShowError("找不到档案");
                    }
                }
            }
        }

        private void btn_toWord_Click(object sender, EventArgs e)
        {
            UIWaitFormService.ShowWaitForm("正在导出，请稍后......");
            if (txtID.Text == "")
            {
                UINotifier.Show("没有档案，请先打开一份档案", UINotifierType.WARNING);
                UIWaitFormService.HideWaitForm();
                return;
            }
            string strName = txtID.Text + "职工基本信息表（" + txtName.Text + "）";
            object path = "D:\\" + strName + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + ".doc";
            object Nothing = Missing.Value;
            MSWord.Application wordAPP = new MSWord.Application();
            MSWord.Document wordDoc = wordAPP.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }
            
            MSWord.Table table = wordDoc.Tables.Add(wordAPP.Selection.Range, 14, 6, ref Nothing, ref Nothing);
            table.Rows.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAtLeast;
            table.Rows.Height = wordAPP.CentimetersToPoints(float.Parse("0.8"));
            table.Range.Font.Size = 10;
            table.Range.Font.Name = "宋体";
            table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleDouble;
            table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            table.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalBottom;

            //合并单元格
            table.Cell(1, 5).Merge(table.Cell(5, 6));
            table.Cell(6, 5).Merge(table.Cell(6, 6));
            table.Cell(9, 4).Merge(table.Cell(9, 6));
            table.Cell(12, 2).Merge(table.Cell(12, 6));
            table.Cell(13, 2).Merge(table.Cell(13, 6));
            table.Cell(14, 2).Merge(table.Cell(14, 6));

            //赋值开始
            table.Cell(1, 1).Range.Text = "职工编号";
            table.Cell(1, 2).Range.Text = txtID.Text;
            table.Cell(1, 3).Range.Text = "职工姓名";
            table.Cell(1, 4).Range.Text = txtName.Text;

            bool isPic = false;
            if (picPhoto != null)
            {
                isPic = true;
                picPhoto.Image.Save(@"D:\" + "职工基本信息表（" + txtName.Text + "）" + ".bmp");
            }
            if (isPic)
            {
                table.Cell(1, 5).Range.InlineShapes.AddPicture(@"D:\" + "职工基本信息表（" + txtName.Text + "）" + ".bmp", false, true, table.Cell(1, 5).Range);
                File.Delete(@"D:\" + "职工基本信息表（" + txtName.Text + "）" + ".bmp");
            }
            else
            {
                table.Cell(1, 5).Range.Text = "缺少照片";
            }

            table.Cell(2, 1).Range.Text ="民族类别";
            table.Cell(2, 2).Range.Text =comFolk.Text;
            table.Cell(2, 3).Range.Text ="出生日期";
            table.Cell(2, 4).Range.Text =uiDatePicker1.Text;

            table.Cell(3, 1).Range.Text ="年龄";
            table.Cell(3, 2).Range.Text =txtAge.Text;
            table.Cell(3, 3).Range.Text ="文化程度";
            table.Cell(3, 4).Range.Text =comCul.Text;

            table.Cell(4, 1).Range.Text ="婚姻";
            table.Cell(4, 2).Range.Text =comMarr.Text;
            table.Cell(4, 3).Range.Text ="性别";
            table.Cell(4, 4).Range.Text =comSex.Text;

            table.Cell(5, 1).Range.Text ="政治面貌";
            table.Cell(5, 2).Range.Text =comVisage.Text;
            table.Cell(5, 3).Range.Text ="单位工作时间";
            table.Cell(5, 4).Range.Text =uiDatePicker2.Text;

            table.Cell(6, 1).Range.Text ="籍贯";
            table.Cell(6, 2).Range.Text =txtBeAware.Text;
            table.Cell(6, 3).Range.Text =txtCity.Text;
            table.Cell(6, 4).Range.Text ="身份证";
            table.Cell(6, 5).Range.Text =txtIDCard.Text;

            table.Cell(7, 1).Range.Text ="工龄";
            table.Cell(7, 2).Range.Text =txtWorkLength.Text;
            table.Cell(7, 3).Range.Text ="职工类别";
            table.Cell(7, 4).Range.Text =comEmployee.Text;
            table.Cell(7, 5).Range.Text ="职务类别";
            table.Cell(7, 6).Range.Text =comBusiness.Text;

            table.Cell(8, 1).Range.Text ="工资类别";
            table.Cell(8, 2).Range.Text =comLaborage.Text;
            table.Cell(8, 3).Range.Text ="部门类别";
            table.Cell(8, 4).Range.Text =comBranch.Text;
            table.Cell(8, 5).Range.Text ="职称类型";
            table.Cell(8, 6).Range.Text =comDuthCall.Text;

            table.Cell(9, 1).Range.Text ="月工资";
            table.Cell(9, 2).Range.Text =txtPay.Text;
            table.Cell(9, 3).Range.Text ="银行账号";
            table.Cell(9, 4).Range.Text =txtBank.Text;

            table.Cell(10, 1).Range.Text ="合同起始日期";
            table.Cell(10, 2).Range.Text =uiDatePicker3.Text;
            table.Cell(10, 3).Range.Text ="合同结束日期";
            table.Cell(10, 4).Range.Text =uiDatePicker4.Text;
            table.Cell(10, 5).Range.Text ="合同年限";
            table.Cell(10, 6).Range.Text =txtY.Text;

            table.Cell(11, 1).Range.Text ="电话";
            table.Cell(11, 2).Range.Text =txtPhone.Text;
            table.Cell(11, 3).Range.Text ="手机";
            table.Cell(11, 4).Range.Text =txtHandset.Text;
            table.Cell(11, 5).Range.Text ="毕业时间";
            table.Cell(11, 6).Range.Text =uiDatePicker5.Text;

            table.Cell(12, 1).Range.Text ="毕业学校";
            table.Cell(12, 2).Range.Text =txtSchool.Text;

            table.Cell(13, 1).Range.Text ="主修专业";
            table.Cell(13, 2).Range.Text =txtSpeciality.Text;

            table.Cell(14, 1).Range.Text ="家庭地址";
            table.Cell(14, 2).Range.Text =txtAddress.Text;



            wordDoc.SaveAs2(ref path, WdSaveFormat.wdFormatDocument, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordAPP.Quit(ref Nothing, ref Nothing, ref Nothing);
            UIMessageTip.ShowOk(path + ".doc文档创建完毕");
            UIWaitFormService.HideWaitForm();
        }

        private void btn_toExcel_Click(object sender, EventArgs e)
        {
            UIWaitFormService.ShowWaitForm("正在导出，请稍后......");
            if (txtID.Text == "")
            {
                UINotifier.Show("没有档案，请先打开一份档案", UINotifierType.WARNING);
                UIWaitFormService.HideWaitForm();
                return;
            }
            string strName = txtID.Text + "职工基本信息表（" + txtName.Text + "）";
            string path = "D:\\" + strName + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + ".xls";
            object Nothing = Missing.Value;
            MSExcel.Application excelApp = new MSExcel.Application();
            MSExcel.Workbook excelWB = excelApp.Application.Workbooks.Add(MSExcel.XlWBATemplate.xlWBATWorksheet);
            Worksheet excelSheet = (MSExcel.Worksheet)(excelWB.Worksheets[1]);

            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }

            MSExcel.Range range= excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[14, 6]];
            range.ColumnWidth = 15;
            range.RowHeight = 25;
            range.Borders.LineStyle = 1;
            range.Font.Size = 12;
            range.Font.Name = "宋体";

            //合并单元格
            MyModule module = new MyModule();
            module.Merge(excelSheet, 1, 5, 5, 6);
            module.Merge(excelSheet, 6, 5, 6, 6);
            module.Merge(excelSheet, 9, 4, 9, 6);
            module.Merge(excelSheet, 12, 2, 12, 6);
            module.Merge(excelSheet, 13, 2, 13, 6);
            module.Merge(excelSheet, 14, 2, 14, 6);

            //开始赋值
            excelApp.Cells[1, 1] = "职工编号";
            excelApp.Cells[1, 2] = txtID.Text;
            excelApp.Cells[1, 3] = "职工姓名";
            excelApp.Cells[1, 4] = txtName.Text;

            bool isPic = false;
            if (picPhoto != null)
            {
                isPic = true;
                picPhoto.Image.Save(@"D:\" + "职工基本信息表（" + txtName.Text + "）" + ".bmp");
            }
            if (isPic)
            {
                excelSheet.Shapes.AddPicture(@"D:\" + "职工基本信息表（" + txtName.Text + "）" + ".bmp", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,360,8,100,115);
                File.Delete(@"D:\" + "职工基本信息表（" + txtName.Text + "）" + ".bmp");
            }
            else
            {
                excelApp.Cells[1,5] = "缺少照片";
            }

            excelApp.Cells[2, 1] = "民族类别";
            excelApp.Cells[2, 2] = comFolk.Text;
            excelApp.Cells[2, 3] = "出生日期";
            excelApp.Cells[2, 4] = uiDatePicker1.Text;
           
            excelApp.Cells[3, 1] = "年龄";
            excelApp.Cells[3, 2] = txtAge.Text;
            excelApp.Cells[3, 3] = "文化程度";
            excelApp.Cells[3, 4] = comCul.Text;
            
            excelApp.Cells[4, 1] = "婚姻";
            excelApp.Cells[4, 2] = comMarr.Text;
            excelApp.Cells[4, 3] = "性别";
            excelApp.Cells[4, 4] = comSex.Text;
            
            excelApp.Cells[5, 1] = "政治面貌";
            excelApp.Cells[5, 2] = comVisage.Text;
            excelApp.Cells[5, 3] = "单位工作时间";
            excelApp.Cells[5, 4] = uiDatePicker2.Text;
          
            excelApp.Cells[6, 1] = "籍贯";
            excelApp.Cells[6, 2] = txtBeAware.Text;
            excelApp.Cells[6, 3] = txtCity.Text;
            excelApp.Cells[6, 4] = "身份证";
            excelApp.Cells[6, 5] = txtIDCard.Text;
         
            excelApp.Cells[7, 1] = "工龄";
            excelApp.Cells[7, 2] = txtWorkLength.Text;
            excelApp.Cells[7, 3] = "职工类别";
            excelApp.Cells[7, 4] = comEmployee.Text;
            excelApp.Cells[7, 5] = "职务类别";
            excelApp.Cells[7, 6] = comBusiness.Text;
         
            excelApp.Cells[8, 1] = "工资类别";
            excelApp.Cells[8, 2] = comLaborage.Text;
            excelApp.Cells[8, 3] = "部门类别";
            excelApp.Cells[8, 4] = comBranch.Text;
            excelApp.Cells[8, 5] = "职称类型";
            excelApp.Cells[8, 6] = comDuthCall.Text;
      
            excelApp.Cells[9, 1] = "月工资";
            excelApp.Cells[9, 2] = txtPay.Text;
            excelApp.Cells[9, 3] = "银行账号";
            excelApp.Cells[9, 4] = txtBank.Text;
         
            excelApp.Cells[10, 1]  = "合同起始日期";
            excelApp.Cells[10, 2]  = uiDatePicker3.Text;
            excelApp.Cells[10, 3]  = "合同结束日期";
            excelApp.Cells[10, 4]  = uiDatePicker4.Text;
            excelApp.Cells[10, 5]  = "合同年限";
            excelApp.Cells[10, 6]  = txtY.Text;
        
            excelApp.Cells[11, 1]  = "电话";
            excelApp.Cells[11, 2]  = txtPhone.Text;
            excelApp.Cells[11, 3]  = "手机";
            excelApp.Cells[11, 4]  = txtHandset.Text;
            excelApp.Cells[11, 5]  = "毕业时间";
            excelApp.Cells[11, 6]  = uiDatePicker5.Text;
        
            excelApp.Cells[12, 1]  = "毕业学校";
            excelApp.Cells[12, 2]  = txtSchool.Text;
     
            excelApp.Cells[13, 1]  = "主修专业";
            excelApp.Cells[13, 2]  = txtSpeciality.Text;
          
            excelApp.Cells[14, 1]  = "家庭地址";
            excelApp.Cells[14, 2]  = txtAddress.Text;

            excelSheet.SaveAs(path,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelWB.Save();
            excelWB.Close(false, Type.Missing, Type.Missing);
            UIMessageTip.ShowOk(path + ".doc文档创建完毕");
            UIWaitFormService.HideWaitForm();
        }

        private void F_FileMan_FormClosed(object sender, FormClosedEventArgs e)
        {
            MyModule.mainFormCon = true;
        }
    }
}

