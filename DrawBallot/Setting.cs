using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace DrawBallot
{
    public partial class Setting : Form
    {
        public event EventHandler PictureChoosen;
        public event EventHandler MusicChoosen;
        DataTable mDataTable = new DataTable();
        int Mode = 0;
        public BindingList<string> listPrize = new BindingList<string>();
        public Setting()
        {

            InitializeComponent();
            boxID.Visible = false;
            boxFN.Visible = false;
            boxLN.Visible = false;
            bCon1.Visible = false;
            bCan1.Visible = false;

            boxPrize.Visible = false;
            bCon2.Visible = false;
            bCan2.Visible = false;

       
            listBox1.DataSource = listPrize;
            

            mDataTable.Columns.Add("ID", typeof(string));
            mDataTable.Columns.Add("First Name", typeof(string));
            mDataTable.Columns.Add("Last Name", typeof(string));
            mDataTable.Columns.Add("isDrawed", typeof(bool));
            mDataTable.Columns["isDrawed"].DefaultValue = false;
            dataGridView1.DataSource = mDataTable;

            if (File.Exists("Participant.csv"))
            {
                mDataTable = Methods.ConvertCSVtoDataTable1("Participant.csv");
            }
            else
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "Choose data source";
                ofd.Filter = "Excel File (*.xlsx)|*.xlsx|07-2003 Excel File (*.xls)|*.xls|all file (*.*)|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText("Participant.csv", Methods.ExcelToCSV(ofd.FileName, ';').ToString());
                    mDataTable = Methods.ConvertCSVtoDataTable1("Participant.csv");
                    MessageBox.Show("Loading Done");
                }
            }
            if (File.Exists("Prizes.csv"))
            {
                List<string> list = new List<string>();
                StreamReader sr = new StreamReader("Prizes.csv");
                foreach(string item in sr.ReadLine().ToString().Split(';'))
                {
                    list.Add(item);
                }
                foreach (string item in list)
                {
                    listPrize.Add(item);
                }
            }
        }

        private void Setting_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = mDataTable;
        }

        private void bLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel File (*.xlsx)|*.xlsx|07-2003 Excel File (*.xls)|*.xls|all file (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText("Participant.csv", Methods.ExcelToCSV(ofd.FileName, ';').ToString());
                mDataTable = Methods.ConvertCSVtoDataTable1("Participant.csv");
                MessageBox.Show("Loading Done");
                dataGridView1.DataSource = mDataTable;
            }
        }

        private void bAdd_Click(object sender, EventArgs e)
        {
            boxID.Text = null;
            boxFN.Text = null;
            boxLN.Text = null;
            boxID.Visible = true;
            boxFN.Visible = true;
            boxLN.Visible = true;
            bAdd1.Enabled = false;
            bDel1.Enabled = false;
            bMod1.Enabled = false;
            bLoad1.Enabled = false;
            bCon1.Visible = true;
            bCan1.Visible = true;
            Mode = 1;
        }

        private void bDel_Click(object sender, EventArgs e)
        {
            int selectedRow;
            selectedRow = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows.RemoveAt(selectedRow);
        }

        private void bMod_Click(object sender, EventArgs e)
        {
            int selectedRow;
            selectedRow = dataGridView1.CurrentCell.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[selectedRow];
            string[] rows = new string[3];
            rows = row.Cells["ID"].Value.ToString().Split('-');
            List<string> arrL = rows.ToList<string>();
            switch (arrL.Count)
            {
                case 1:
                    arrL.Add("");
                    arrL.Add("");
                    break;
                case 2:
                    arrL.Add("");
                    break;
                default:
                    break;
            }
            boxID.Text = arrL[0];
            boxFN.Text = arrL[1];
            boxLN.Text = arrL[2];
            boxID.Visible = true;
            boxFN.Visible = true;
            boxLN.Visible = true;
            bAdd1.Enabled = false;
            bDel1.Enabled = false;
            bMod1.Enabled = false;
            bLoad1.Enabled = false;
            bCon1.Visible = true;
            bCan1.Visible = true;
            Mode = 2;
        }

        private void bReset_Click(object sender, EventArgs e)
        {
            mDataTable.Clear();
        }

        private void bCon_Click(object sender, EventArgs e)
        {
            if (Mode == 1)
            {
                var newrow = mDataTable.NewRow();
                newrow[0] = boxID.Text+"-"+ boxFN.Text + "-" + boxLN.Text;
                newrow[1] = false;
                
                mDataTable.Rows.Add(newrow);
            }
            else
            {
                int selectedRow;
                selectedRow = dataGridView1.CurrentCell.RowIndex;
                DataGridViewRow row = dataGridView1.Rows[selectedRow];
                row.Cells[0].Value = boxID.Text + "-" + boxFN.Text + "-" + boxLN.Text;
                dataGridView1.EndEdit();
                ((DataRowView)dataGridView1.CurrentRow.DataBoundItem).EndEdit();
            }

            boxID.Visible = false;
            boxFN.Visible = false;
            boxLN.Visible = false;
            bAdd1.Enabled = true;
            bDel1.Enabled = true;
            bMod1.Enabled = true;
            bLoad1.Enabled = true;
            bCon1.Visible = false;
            bCan1.Visible = false;
            //mAdapter.Update(mDataTable);
        }

        private void bCan_Click(object sender, EventArgs e)
        {
            boxID.Visible = false;
            boxFN.Visible = false;
            boxLN.Visible = false;
            bAdd1.Enabled = true;
            bDel1.Enabled = true;
            bMod1.Enabled = true;
            bLoad1.Enabled = true;
            bCon1.Visible = false;
            bCan1.Visible = false;
            Mode = 0;
        }

        private void bAdd2_Click(object sender, EventArgs e)
        {
            
            bAdd2.Enabled = false;
            bDel2.Enabled = false;
            bMod2.Enabled = false;
            bReset2.Enabled = false;

            boxPrize.Visible = true;
            bCon2.Visible = true;
            bCan2.Visible = true;
            Mode = 1;
        }

        private void bMod2_Click(object sender, EventArgs e)
        {
            bAdd2.Enabled = false;
            bDel2.Enabled = false;
            bMod2.Enabled = false;
            bReset2.Enabled = false;

            boxPrize.Visible = true;
            bCon2.Visible = true;
            bCan2.Visible = true;
            Mode = 2;
        }

        private void bCon2_Click(object sender, EventArgs e)
        {
            if (Mode == 1)
            {
                listPrize.Add(boxPrize.Text);
            }
            else
            {
                listPrize[listBox1.SelectedIndex]= boxPrize.Text;
            }
            bAdd2.Enabled = true;
            bDel2.Enabled = true;
            bMod2.Enabled = true;
            bReset2.Enabled = true;

            boxPrize.Visible = false;
            bCon2.Visible = false;
            bCan2.Visible = false;
            Mode = 0;
        }

        private void bBrowsePic_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            try
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = ofd.FileName;
                    PictureChoosen(sender,e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void bCan2_Click(object sender, EventArgs e)
        {
            bAdd2.Enabled = true;
            bDel2.Enabled = true;
            bMod2.Enabled = true;
            bReset2.Enabled = true;

            boxPrize.Visible = false;
            bCon2.Visible = false;
            bCan2.Visible = false;
            Mode = 0;
        }

        private void bReset2_Click(object sender, EventArgs e)
        {
            listPrize.Clear();
            listPrize.Add("First Place");
            listPrize.Add("Second Place");
            listPrize.Add("Third Place");
        }

        private void bUp_Click(object sender, EventArgs e)
        {
            int n = listBox1.SelectedIndex;
            string temp = listPrize[n];
            if (n >0)
            {
                listPrize.RemoveAt(n);
                listPrize.Insert(n - 1, temp);
                listBox1.SetSelected(n - 1,true);
            }
        }

        private void bDown_Click(object sender, EventArgs e)
        {
            int n = listBox1.SelectedIndex;
            string temp = listPrize[n];
            if (n < listPrize.Count - 1)
            {
                listPrize.RemoveAt(n);
                listPrize.Insert(n + 1, temp);
                listBox1.SetSelected(n + 1, true);
            }
        }

        private void bDel2_Click(object sender, EventArgs e)
        {
            listPrize.RemoveAt(listBox1.SelectedIndex);
        }

        public DataTable Table
        {
            get
            {
                return this.mDataTable;
            }
            set
            {
                this.mDataTable = value;
            }
        }

        public string Picture
        {
            get
            {
                return textBox2.Text;
            }
            set
            {

            }
        }

        public string Music
        {
            get
            {
                return textBox3.Text;
            }
            set
            {

            }
        }

        private void bBrowseSound_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            try
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    textBox3.Text = ofd.FileName;
                    MusicChoosen(sender, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
