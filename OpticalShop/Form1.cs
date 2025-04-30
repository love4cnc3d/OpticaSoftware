namespace OpticalShop
{
    using ADOX;
    using System;
    using System.ComponentModel;
    using System.Data;
    using System.Data.OleDb;
    using System.Drawing;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public class Form1 : Form
    {
        private OleDbConnection conn;
        private string connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/OpticalShop/Database/OpticalDb.mdb;Jet OLEDB:Engine Type=5";
        private IContainer components;
        private Label label20;
        private ComboBox lens3;
        private Label label19;
        private ComboBox lens2;
        private Label label18;
        private Label label17;
        private Label label16;
        private Label label15;
        private GroupBox groupBox1;
        private Label label14;
        private Label label13;
        private Label label12;
        private Label label11;
        private Label label10;
        private Label label9;
        private TextBox t7;
        private Button button1;
        private Label label8;
        private TextBox t6;
        private Label label7;
        private RadioButton r2;
        private TextBox t5;
        private TextBox t3;
        private Label label3;
        private Label label2;
        private TabPage tabPage1;
        private RadioButton r1;
        private Label label6;
        private Label label5;
        private Label label4;
        private TextBox t2;
        private TabPage tabPage2;
        private TabControl tabControl1;
        private Button button3;
        private Button button2;
        private Label label23;
        private TextBox f2;
        private Label label22;
        private ComboBox f1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem editToolStripMenuItem;
        private ToolStripMenuItem toolToolStripMenuItem;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripMenuItem helpToolStripMenuItem1;
        private ToolStripMenuItem aboutToolStripMenuItem;
        private ToolStripMenuItem saveAsToolStripMenuItem;
        private ToolStripMenuItem excelSheetToolStripMenuItem;
        private ToolStripMenuItem accessFIleToolStripMenuItem;
        private ToolStripMenuItem sendToToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem1;
        private ToolStripMenuItem backupToolStripMenuItem;
        private ToolStripMenuItem exportDatabaseToolStripMenuItem;
        private ToolStripMenuItem importDatabaseToolStripMenuItem;
        private DateTimePicker d4;
        private GroupBox groupBox3;
        private DataGridView dataGridView1;
        private Label label29;
        private GroupBox groupBox4;
        private Label lb_id;
        private Label label41;
        private GroupBox groupBox5;
        private Label lb_contact;
        private Label lb_name;
        private Label label30;
        private Label label31;
        private Label cust_id;
        private Label label33;
        private Label label32;
        private SaveFileDialog saveFileDialog1;
        private OpenFileDialog openFileDialog1;
        private Label label43;
        private Label label42;
        private Label label40;
        private Label label21;
        private TextBox lens4;
        private ComboBox Search_by;
        private TextBox Search_txt;
        private Label label44;
        private ComboBox left2;
        private ComboBox right2;
        private ComboBox left1;
        private ComboBox right1;
        private ComboBox left3;
        private ComboBox right3;
        private Label label25;
        private Label p5;
        private TextBox p2;
        private Label label26;
        private Label label28;
        private TextBox p3;
        private Label label27;
        private GroupBox groupBox2;
        private DataGridView dataGridView2;
        private Button button4;
        private TextBox p1;
        private Label label24;
        private TextBox p4;
        private Label label46;
        private Label label45;
        private TextBox Pre_id;
        private Button button5;
        private Button button6;
        private Label pen_status;
        private GroupBox groupBox6;
        private Button ll6;
        private Button lr6;
        private Button ll5;
        private Button lr5;
        private Button ll4;
        private Button lr4;
        private Button ll3;
        private Button lr3;
        private Button ll2;
        private Button lr2;
        private Button ll1;
        private Button lr1;
        private Label fm_type;
        private Label lb_side;
        private Label lb_for;
        private Label lb_type;
        private Label label58;
        private Label label59;
        private Label label60;
        private Label label62;
        private Label label49;
        private Label label48;
        private Label label50;
        private Label label34;
        private Label label39;
        private Label label35;
        private Label label38;
        private Label label36;
        private Label label37;
        private Label label47;
        private TextBox pre_id2;
        private Button button7;
        private Button button8;
        private ComboBox left4;
        private ComboBox right4;
        private ComboBox lens1;
        private Label label51;
        private Button total3;
        private Label label53;
        private Button total2;
        private Label label52;
        private Button total1;
        private ComboBox left_6;
        private ComboBox right_6;
        private ComboBox left_5;
        private ComboBox right_5;
        private Button button9;
        private Label label1;
        private Button button10;

        public Form1()
        {
            this.InitializeComponent();
            base.WindowState = FormWindowState.Maximized;
            if (!Directory.Exists("C:/OpticalShop/Backups/"))
            {
                Directory.CreateDirectory("C:/OpticalShop/Backups/");
                MessageBox.Show("Directory Created!! @Location:C:/OpticalShop/Backups/");
            }
            if (!Directory.Exists("C:/OpticalShop/Database/"))
            {
                Directory.CreateDirectory("C:/OpticalShop/Database/");
                MessageBox.Show("Directory Created!! @Location:C:/OpticalShop/Database/ ");
            }
            if (!File.Exists("C:/OpticalShop/Database/OpticalDb.mdb"))
            {
                Catalog catalog = (Catalog) Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00000602-0000-0010-8000-00AA006D2EA4")));
                Table item = (Table) Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00000609-0000-0010-8000-00AA006D2EA4")));
                item.Name = "Login";
                item[].Append("Username", DataTypeEnum.adLongVarWChar, 0xff);
                item[].Append("password", DataTypeEnum.adLongVarWChar, 0xff);
                try
                {
                    catalog.Create(this.connstr);
                    catalog[].Append(item);
                }
                catch (Exception exception1)
                {
                    MessageBox.Show(exception1.Message);
                }
                try
                {
                    this.conn = new OleDbConnection();
                    this.conn.ConnectionString = this.connstr;
                    this.conn.Open();
                    new OleDbCommand("CREATE TABLE Customer (`ID` IDENTITY(1,1),`Ref_no` varchar(25), `Ex_by`  varchar(25),`name`  varchar(25),`birthdate`  varchar(25), `Addr`  varchar(50), `Gender` varchar(25), `Contact` varchar(25), `Age` varchar(25),`Order_date` date,PRIMARY KEY(ID));", this.conn).ExecuteNonQuery();
                    this.conn.Close();
                    this.conn.Open();
                    new OleDbCommand("CREATE TABLE Lens_Pre (Pre_ID IDENTITY(1,1),ID number, r_sph varchar(25),r_cyl varchar(25),r_axis varchar(25),r_vd varchar(25),r_near varchar(25), r_vn varchar(25),l_sph varchar(25),l_cyl varchar(25),l_axis varchar(25),l_vd varchar(25),l_near varchar(25), l_vn varchar(25),ls_type varchar(25),ls_for varchar(25), ls_side varchar(25),ls_price number, fm_type varchar(25), fm_price number, `Paid_Amt` varchar(25), `total_Amt`  varchar(25), `Pending_Status`  varchar(25), `Order_date` date,`Extra_disc` varchar(25),`Extra_charges` varchar(25),PRIMARY KEY(Pre_ID));", this.conn).ExecuteNonQuery();
                    this.conn.Close();
                }
                catch (OleDbException exception3)
                {
                    MessageBox.Show(exception3.Message);
                }
                MessageBox.Show("Database Created!!");
            }
            this.Refresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.t3.Text.Length == 0)
            {
                MessageBox.Show("Please Enter the name!");
            }
            else
            {
                try
                {
                    string str = "";
                    str = !this.r1.Checked ? "Female" : "Male";
                    this.conn = new OleDbConnection();
                    this.conn.ConnectionString = this.connstr;
                    this.conn.Open();
                    string[] strArray = new string[] { "INSERT INTO `Customer` (`Ref_no`, `Ex_by`,`name`,`birthdate`, `Addr`, `Gender`, `Contact`, `Age`,`Order_date` ) VALUES ('--','", this.t2.Text, "','", this.t3.Text, "','", this.d4.Text, "', '", this.t5.Text, "', '" };
                    strArray[9] = str;
                    strArray[10] = "','";
                    strArray[11] = this.t6.Text;
                    strArray[12] = "','";
                    strArray[13] = this.t7.Text;
                    strArray[14] = "',NOW());";
                    new OleDbCommand(string.Concat(strArray), this.conn).ExecuteNonQuery();
                    OleDbCommand command2 = new OleDbCommand("SELECT ID,name,Contact from Customer ORDER BY ID DESC;", this.conn);
                    try
                    {
                        OleDbDataReader reader = command2.ExecuteReader();
                        if (reader.Read())
                        {
                            reader[0].ToString();
                            this.lb_name.Text = reader["name"].ToString();
                            this.lb_contact.Text = reader["Contact"].ToString();
                        }
                    }
                    catch (Exception exception1)
                    {
                        MessageBox.Show(exception1.Message);
                    }
                    this.Refresh();
                    this.conn.Close();
                    MessageBox.Show("Inserted!!");
                }
                catch (OleDbException exception3)
                {
                    MessageBox.Show(exception3.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string str;
            string str2;
            string str3;
            this.t7.Text = str = "";
            this.t6.Text = str2 = str;
            this.t5.Text = str3 = str2;
            this.t2.Text = this.t3.Text = str3;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                this.conn = new OleDbConnection();
                this.conn.ConnectionString = this.connstr;
                string str = "";
                str = !this.r1.Checked ? "Female" : "Male";
                this.conn.Open();
                string[] strArray = new string[] { "UPDATE Customer set `Ref_no`='--', `Ex_by`='", this.t2.Text, "',`name`='", this.t3.Text, "', `Addr`='", this.t5.Text, "', `Gender`='", str, "', `Contact`='" };
                strArray[9] = this.t6.Text;
                strArray[10] = "', `Age`='";
                strArray[11] = this.t7.Text;
                strArray[12] = "'  where ID=";
                strArray[13] = this.cust_id.Text;
                new OleDbCommand(string.Concat(strArray), this.conn).ExecuteNonQuery();
                MessageBox.Show("Updated!!!");
                this.conn.Close();
            }
            catch (OleDbException exception1)
            {
                MessageBox.Show(exception1.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            this.conn = new OleDbConnection();
            this.conn.ConnectionString = this.connstr;
            this.conn.Open();
            string[] strArray = new string[] { "INSERT INTO `Lens_Pre` (ID, r_sph ,r_cyl ,r_axis ,r_vd ,r_near , r_vn ,l_sph ,l_cyl ,l_axis ,l_vd ,l_near , l_vn ,ls_type ,ls_for , ls_side ,ls_price, fm_type , fm_price,`Paid_Amt`, `total_Amt`, `Pending_Status`, `Order_date`, `Extra_disc`, `Extra_charges`) VALUES (", this.cust_id.Text, ",'", this.right1.Text, "','", this.right2.Text, "','", this.right3.Text, "','" };
            strArray[9] = this.right4.Text;
            strArray[10] = "','";
            strArray[11] = this.right_5.Text;
            strArray[12] = "','";
            strArray[13] = this.right_6.Text;
            strArray[14] = "','";
            strArray[15] = this.left1.Text;
            strArray[0x10] = "','";
            strArray[0x11] = this.left2.Text;
            strArray[0x12] = "','";
            strArray[0x13] = this.left3.Text;
            strArray[20] = "','";
            strArray[0x15] = this.left4.Text;
            strArray[0x16] = "','";
            strArray[0x17] = this.left_5.Text;
            strArray[0x18] = "','";
            strArray[0x19] = this.left_6.Text;
            strArray[0x1a] = "','";
            strArray[0x1b] = this.lens1.Text;
            strArray[0x1c] = "','";
            strArray[0x1d] = this.lens2.Text;
            strArray[30] = "','";
            strArray[0x1f] = this.lens3.Text;
            strArray[0x20] = "',";
            strArray[0x21] = this.lens4.Text;
            strArray[0x22] = ",'";
            strArray[0x23] = this.f1.Text;
            strArray[0x24] = "',";
            strArray[0x25] = this.f2.Text;
            strArray[0x26] = ",'";
            strArray[0x27] = this.p3.Text;
            strArray[40] = "','";
            strArray[0x29] = this.p4.Text;
            strArray[0x2a] = "','";
            strArray[0x2b] = this.pen_status.Text;
            strArray[0x2c] = "',NOW(),'";
            strArray[0x2d] = this.p1.Text;
            strArray[0x2e] = "','";
            strArray[0x2f] = this.p2.Text;
            strArray[0x30] = "');";
            new OleDbCommand(string.Concat(strArray), this.conn).ExecuteNonQuery();
            OleDbCommand command2 = new OleDbCommand("SELECT * from lens_Pre where ID=" + this.cust_id.Text, this.conn);
            try
            {
                OleDbDataReader reader = command2.ExecuteReader();
                if (reader.Read())
                {
                    this.lb_id.Text = reader["ID"].ToString();
                    this.lr1.Text = reader[1].ToString();
                    this.lr2.Text = reader[2].ToString();
                    this.lr3.Text = reader[3].ToString();
                    this.lr4.Text = reader[4].ToString();
                    this.lr5.Text = reader[5].ToString();
                    this.lr6.Text = reader[6].ToString();
                    this.ll1.Text = reader[7].ToString();
                    this.ll2.Text = reader[8].ToString();
                    this.ll3.Text = reader[9].ToString();
                    this.ll4.Text = reader[10].ToString();
                    this.ll5.Text = reader[11].ToString();
                    this.ll6.Text = reader[12].ToString();
                    this.lb_type.Text = reader["ls_type"].ToString();
                    this.lb_for.Text = reader["ls_for"].ToString();
                    this.lb_side.Text = reader["ls_side"].ToString();
                    this.fm_type.Text = reader["fm_type"].ToString();
                }
            }
            catch (Exception exception1)
            {
                MessageBox.Show(exception1.Message);
            }
            this.Refresh_Pre();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            try
            {
                this.conn = new OleDbConnection();
                this.conn.ConnectionString = this.connstr;
                this.conn.Open();
                string[] strArray = new string[] { "UPDATE Lens_Pre set r_sph='", this.right1.Text, "' ,r_cyl='", this.right2.Text, "' ,r_axis='", this.right3.Text, "' ,r_vd='", this.right4.Text, "' ,r_near='" };
                strArray[9] = this.right_5.Text;
                strArray[10] = "', r_vn='";
                strArray[11] = this.right_6.Text;
                strArray[12] = "' ,l_sph='";
                strArray[13] = this.left1.Text;
                strArray[14] = "' ,l_cyl='";
                strArray[15] = this.left2.Text;
                strArray[0x10] = "' ,l_axis='";
                strArray[0x11] = this.left3.Text;
                strArray[0x12] = "' ,l_vd='";
                strArray[0x13] = this.left4.Text;
                strArray[20] = "' ,l_near='";
                strArray[0x15] = this.left_5.Text;
                strArray[0x16] = "' , l_vn='";
                strArray[0x17] = this.left_6.Text;
                strArray[0x18] = "' ,ls_type='";
                strArray[0x19] = this.lens1.Text;
                strArray[0x1a] = "' ,ls_for='";
                strArray[0x1b] = this.lens2.Text;
                strArray[0x1c] = "' , ls_side='";
                strArray[0x1d] = this.lens3.Text;
                strArray[30] = "' ,ls_price='";
                strArray[0x1f] = this.lens4.Text;
                strArray[0x20] = "', fm_type='";
                strArray[0x21] = this.f1.Text;
                strArray[0x22] = "' , fm_price='";
                strArray[0x23] = this.f2.Text;
                strArray[0x24] = "',`Paid_Amt`='";
                strArray[0x25] = this.p3.Text;
                strArray[0x26] = "', `total_Amt`='";
                strArray[0x27] = this.p4.Text;
                strArray[40] = "', `Pending_Status`='";
                strArray[0x29] = this.pen_status.Text;
                strArray[0x2a] = "', `Extra_disc`='";
                strArray[0x2b] = this.p1.Text;
                strArray[0x2c] = "', `Extra_charges`='";
                strArray[0x2d] = this.p2.Text;
                strArray[0x2e] = "' where ID=";
                strArray[0x2f] = this.cust_id.Text;
                strArray[0x30] = " AND Pre_Id=";
                strArray[0x31] = this.Pre_id.Text;
                new OleDbCommand(string.Concat(strArray), this.conn).ExecuteNonQuery();
                MessageBox.Show("Updated!!!");
                this.conn.Close();
            }
            catch (OleDbException exception1)
            {
                MessageBox.Show(exception1.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Refresh();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                this.conn = new OleDbConnection();
                this.conn.ConnectionString = this.connstr;
                this.conn.Open();
                new OleDbCommand("DELETE FROM Customer WHERE ID=" + this.cust_id.Text, this.conn).ExecuteNonQuery();
                new OleDbCommand("DELETE FROM Lens_Pre WHERE ID=" + this.cust_id.Text, this.conn).ExecuteNonQuery();
                MessageBox.Show("DELETED!!");
                this.Refresh();
                this.conn.Close();
            }
            catch (OleDbException exception1)
            {
                MessageBox.Show(exception1.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                this.conn = new OleDbConnection();
                this.conn.ConnectionString = this.connstr;
                this.conn.Open();
                new OleDbCommand("DELETE FROM Lens_Pre WHERE Pre_ID=" + this.Pre_id.Text, this.conn).ExecuteNonQuery();
                MessageBox.Show("DELETED Prescription!!");
                this.Refresh_Pre();
                this.conn.Close();
            }
            catch (OleDbException exception1)
            {
                MessageBox.Show(exception1.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string str;
            string str2;
            string str3;
            string str4;
            string str6;
            string str7;
            string str8;
            string str9;
            string str11;
            string str12;
            string str14;
            string str15;
            this.right_6.Text = str = "";
            this.right_5.Text = str2 = str;
            this.right4.Text = str3 = str2;
            this.right3.Text = str4 = str3;
            this.right1.Text = this.right2.Text = str4;
            this.left_6.Text = str6 = "";
            this.left_5.Text = str7 = str6;
            this.left4.Text = str8 = str7;
            this.left3.Text = str9 = str8;
            this.left1.Text = this.left2.Text = str9;
            this.Pre_id.Text = str11 = "";
            this.p1.Text = str12 = str11;
            this.lens1.Text = this.f1.Text = str12;
            this.p3.Text = str14 = "0";
            this.p2.Text = str15 = str14;
            this.lens4.Text = this.f2.Text = str15;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            this.tabControl1.SelectedTab = this.tabPage2;
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            this.cust_id.Text = this.lb_id.Text = this.dataGridView1.Rows[e.RowIndex].Cells["ID"].Value.ToString();
            this.t2.Text = this.dataGridView1.Rows[e.RowIndex].Cells["Ex_by"].Value.ToString();
            this.t3.Text = this.lb_name.Text = this.dataGridView1.Rows[e.RowIndex].Cells["name"].Value.ToString();
            this.t5.Text = this.dataGridView1.Rows[e.RowIndex].Cells["Addr"].Value.ToString();
            this.t6.Text = this.lb_contact.Text = this.dataGridView1.Rows[e.RowIndex].Cells["Contact"].Value.ToString();
            this.t7.Text = this.dataGridView1.Rows[e.RowIndex].Cells["Age"].Value.ToString();
            this.Refresh_Pre();
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string str2;
            this.conn = new OleDbConnection();
            this.conn.ConnectionString = this.connstr;
            this.conn.Open();
            this.Pre_id.Text = str2 = this.dataGridView2.Rows[e.RowIndex].Cells["Pre_Id"].Value.ToString();
            string str = this.pre_id2.Text = str2;
            OleDbCommand command = new OleDbCommand("SELECT * from lens_Pre where ID=" + this.cust_id.Text + " AND Pre_Id=" + str, this.conn);
            try
            {
                OleDbDataReader reader = command.ExecuteReader();
                if (reader.Read())
                {
                    string str22;
                    this.lb_id.Text = reader["ID"].ToString();
                    this.right1.Text = this.lr1.Text = reader["r_sph"].ToString();
                    this.right2.Text = this.lr2.Text = reader["r_cyl"].ToString();
                    this.right3.Text = this.lr3.Text = reader["r_AXIS"].ToString();
                    this.right4.Text = this.lr4.Text = reader["r_vd"].ToString();
                    this.right_5.Text = this.lr5.Text = reader["r_near"].ToString();
                    this.right_6.Text = this.lr6.Text = reader["r_vn"].ToString();
                    this.left1.Text = this.ll1.Text = reader["l_sph"].ToString();
                    this.left2.Text = this.ll2.Text = reader["l_cyl"].ToString();
                    this.left3.Text = this.ll3.Text = reader["l_AXIS"].ToString();
                    this.left4.Text = this.ll4.Text = reader["l_vd"].ToString();
                    this.left_5.Text = this.ll5.Text = reader["l_near"].ToString();
                    this.left_6.Text = this.ll6.Text = reader["l_vn"].ToString();
                    this.lens1.SelectedItem = this.lb_type.Text = reader["ls_type"].ToString();
                    this.lens2.SelectedItem = this.lb_for.Text = reader["ls_for"].ToString();
                    this.lens3.SelectedItem = this.lb_side.Text = reader["ls_side"].ToString();
                    this.total1.Text = this.lens4.Text = reader["ls_price"].ToString();
                    this.f1.Text = this.fm_type.Text = reader["fm_type"].ToString();
                    this.total2.Text = this.f2.Text = reader["fm_price"].ToString();
                    this.p3.Text = reader["Paid_Amt"].ToString();
                    this.p1.Text = reader["Extra_disc"].ToString();
                    this.p2.Text = reader["Extra_charges"].ToString();
                    this.p1.Text = str22 = reader["total_Amt"].ToString();
                    this.total3.Text = this.p4.Text = str22;
                }
            }
            catch (Exception exception1)
            {
                MessageBox.Show(exception1.Message);
            }
            this.conn.Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void exportDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.saveFileDialog1.ShowDialog();
        }

        private void f2_TextChanged(object sender, EventArgs e)
        {
            float num = (Convert.ToInt32("0" + this.f2.Text) + Convert.ToInt32("0" + this.lens4.Text)) + Convert.ToInt32("0" + this.p2.Text);
            this.p3.Text = this.p4.Text = num.ToString();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            loading_thread _thread = new loading_thread();
            _thread.Show();
            _thread.Close();
            _thread.Dispose();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            loading_thread _thread = new loading_thread();
            _thread.Show();
            _thread.Close();
            _thread.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void Form1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.S))
            {
                this.button1_Click(null, null);
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }

        private void importDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.ShowDialog();
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(Form1));
            this.label20 = new Label();
            this.lens3 = new ComboBox();
            this.label19 = new Label();
            this.lens2 = new ComboBox();
            this.label18 = new Label();
            this.label17 = new Label();
            this.label16 = new Label();
            this.label15 = new Label();
            this.groupBox1 = new GroupBox();
            this.button9 = new Button();
            this.left_6 = new ComboBox();
            this.right_6 = new ComboBox();
            this.left_5 = new ComboBox();
            this.right_5 = new ComboBox();
            this.left4 = new ComboBox();
            this.right4 = new ComboBox();
            this.button8 = new Button();
            this.pen_status = new Label();
            this.button5 = new Button();
            this.p4 = new TextBox();
            this.p1 = new TextBox();
            this.label24 = new Label();
            this.button4 = new Button();
            this.label25 = new Label();
            this.p5 = new Label();
            this.p2 = new TextBox();
            this.label26 = new Label();
            this.label28 = new Label();
            this.p3 = new TextBox();
            this.label27 = new Label();
            this.left3 = new ComboBox();
            this.right3 = new ComboBox();
            this.left2 = new ComboBox();
            this.right2 = new ComboBox();
            this.left1 = new ComboBox();
            this.right1 = new ComboBox();
            this.label23 = new Label();
            this.label21 = new Label();
            this.f2 = new TextBox();
            this.lens4 = new TextBox();
            this.label22 = new Label();
            this.f1 = new ComboBox();
            this.lens1 = new ComboBox();
            this.label14 = new Label();
            this.label13 = new Label();
            this.label12 = new Label();
            this.label11 = new Label();
            this.label10 = new Label();
            this.label9 = new Label();
            this.t7 = new TextBox();
            this.button1 = new Button();
            this.label8 = new Label();
            this.t6 = new TextBox();
            this.label7 = new Label();
            this.r2 = new RadioButton();
            this.t5 = new TextBox();
            this.t3 = new TextBox();
            this.label3 = new Label();
            this.label2 = new Label();
            this.tabPage1 = new TabPage();
            this.button7 = new Button();
            this.label43 = new Label();
            this.label42 = new Label();
            this.label40 = new Label();
            this.label33 = new Label();
            this.label32 = new Label();
            this.d4 = new DateTimePicker();
            this.button3 = new Button();
            this.button2 = new Button();
            this.r1 = new RadioButton();
            this.label6 = new Label();
            this.label5 = new Label();
            this.label4 = new Label();
            this.t2 = new TextBox();
            this.cust_id = new Label();
            this.label29 = new Label();
            this.tabPage2 = new TabPage();
            this.label46 = new Label();
            this.label45 = new Label();
            this.Pre_id = new TextBox();
            this.groupBox2 = new GroupBox();
            this.dataGridView2 = new DataGridView();
            this.tabControl1 = new TabControl();
            this.menuStrip1 = new MenuStrip();
            this.fileToolStripMenuItem = new ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new ToolStripMenuItem();
            this.excelSheetToolStripMenuItem = new ToolStripMenuItem();
            this.accessFIleToolStripMenuItem = new ToolStripMenuItem();
            this.sendToToolStripMenuItem = new ToolStripMenuItem();
            this.exitToolStripMenuItem = new ToolStripMenuItem();
            this.exitToolStripMenuItem1 = new ToolStripMenuItem();
            this.editToolStripMenuItem = new ToolStripMenuItem();
            this.toolToolStripMenuItem = new ToolStripMenuItem();
            this.backupToolStripMenuItem = new ToolStripMenuItem();
            this.exportDatabaseToolStripMenuItem = new ToolStripMenuItem();
            this.importDatabaseToolStripMenuItem = new ToolStripMenuItem();
            this.helpToolStripMenuItem = new ToolStripMenuItem();
            this.helpToolStripMenuItem1 = new ToolStripMenuItem();
            this.aboutToolStripMenuItem = new ToolStripMenuItem();
            this.groupBox3 = new GroupBox();
            this.button6 = new Button();
            this.Search_by = new ComboBox();
            this.Search_txt = new TextBox();
            this.dataGridView1 = new DataGridView();
            this.label44 = new Label();
            this.groupBox4 = new GroupBox();
            this.label47 = new Label();
            this.pre_id2 = new TextBox();
            this.lb_id = new Label();
            this.label41 = new Label();
            this.groupBox6 = new GroupBox();
            this.button10 = new Button();
            this.label1 = new Label();
            this.label53 = new Label();
            this.total2 = new Button();
            this.label52 = new Label();
            this.total1 = new Button();
            this.label51 = new Label();
            this.total3 = new Button();
            this.ll6 = new Button();
            this.lr6 = new Button();
            this.ll5 = new Button();
            this.lr5 = new Button();
            this.ll4 = new Button();
            this.lr4 = new Button();
            this.ll3 = new Button();
            this.lr3 = new Button();
            this.ll2 = new Button();
            this.lr2 = new Button();
            this.ll1 = new Button();
            this.lr1 = new Button();
            this.fm_type = new Label();
            this.lb_side = new Label();
            this.lb_for = new Label();
            this.lb_type = new Label();
            this.label58 = new Label();
            this.label59 = new Label();
            this.label60 = new Label();
            this.label62 = new Label();
            this.label49 = new Label();
            this.label48 = new Label();
            this.label50 = new Label();
            this.label34 = new Label();
            this.label39 = new Label();
            this.label35 = new Label();
            this.label38 = new Label();
            this.label36 = new Label();
            this.label37 = new Label();
            this.groupBox5 = new GroupBox();
            this.lb_contact = new Label();
            this.lb_name = new Label();
            this.label30 = new Label();
            this.label31 = new Label();
            this.saveFileDialog1 = new SaveFileDialog();
            this.openFileDialog1 = new OpenFileDialog();
            this.groupBox1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((ISupportInitialize) this.dataGridView2).BeginInit();
            this.tabControl1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((ISupportInitialize) this.dataGridView1).BeginInit();
            this.groupBox4.SuspendLayout();
            this.groupBox6.SuspendLayout();
            base.SuspendLayout();
            this.label20.AutoSize = true;
            this.label20.Location = new Point(6, 0x7d);
            this.label20.Name = "label20";
            this.label20.Size = new Size(0x36, 13);
            this.label20.TabIndex = 0x2e;
            this.label20.Text = "Lens Side";
            this.lens3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.lens3.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.lens3.DropDownStyle = ComboBoxStyle.DropDownList;
            this.lens3.FormattingEnabled = true;
            object[] items = new object[] { "Both", "Right", "Left" };
            this.lens3.Items.AddRange(items);
            this.lens3.Location = new Point(80, 0x7a);
            this.lens3.Name = "lens3";
            this.lens3.Size = new Size(100, 0x15);
            this.lens3.TabIndex = 0x16;
            this.label19.AutoSize = true;
            this.label19.Location = new Point(0xde, 0x63);
            this.label19.Name = "label19";
            this.label19.Size = new Size(0x30, 13);
            this.label19.TabIndex = 0x2c;
            this.label19.Text = "Lens For";
            this.lens2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.lens2.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.lens2.DropDownStyle = ComboBoxStyle.DropDownList;
            this.lens2.FormattingEnabled = true;
            object[] objArray2 = new object[] { "Distance ", "Near", "Bifocal" };
            this.lens2.Items.AddRange(objArray2);
            this.lens2.Location = new Point(0x128, 0x60);
            this.lens2.Name = "lens2";
            this.lens2.Size = new Size(100, 0x15);
            this.lens2.TabIndex = 0x15;
            this.label18.AutoSize = true;
            this.label18.Location = new Point(6, 0x63);
            this.label18.Name = "label18";
            this.label18.Size = new Size(0x39, 13);
            this.label18.TabIndex = 0x2a;
            this.label18.Text = "Lens Type";
            this.label17.AutoSize = true;
            this.label17.Location = new Point(0x156, 0x1a);
            this.label17.Name = "label17";
            this.label17.Size = new Size(0x16, 13);
            this.label17.TabIndex = 0x27;
            this.label17.Text = "VN";
            this.label16.AutoSize = true;
            this.label16.Location = new Point(0x121, 0x11);
            this.label16.Name = "label16";
            this.label16.Size = new Size(40, 0x1a);
            this.label16.TabIndex = 0x24;
            this.label16.Text = "NEAR \r\nADD";
            this.label15.AutoSize = true;
            this.label15.Location = new Point(0xee, 0x1a);
            this.label15.Name = "label15";
            this.label15.Size = new Size(0x16, 13);
            this.label15.TabIndex = 0x21;
            this.label15.Text = "VD";
            this.groupBox1.Controls.Add(this.button9);
            this.groupBox1.Controls.Add(this.left_6);
            this.groupBox1.Controls.Add(this.right_6);
            this.groupBox1.Controls.Add(this.left_5);
            this.groupBox1.Controls.Add(this.right_5);
            this.groupBox1.Controls.Add(this.left4);
            this.groupBox1.Controls.Add(this.right4);
            this.groupBox1.Controls.Add(this.button8);
            this.groupBox1.Controls.Add(this.pen_status);
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.p4);
            this.groupBox1.Controls.Add(this.p1);
            this.groupBox1.Controls.Add(this.label24);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.label25);
            this.groupBox1.Controls.Add(this.p5);
            this.groupBox1.Controls.Add(this.p2);
            this.groupBox1.Controls.Add(this.label26);
            this.groupBox1.Controls.Add(this.label28);
            this.groupBox1.Controls.Add(this.p3);
            this.groupBox1.Controls.Add(this.label27);
            this.groupBox1.Controls.Add(this.left3);
            this.groupBox1.Controls.Add(this.right3);
            this.groupBox1.Controls.Add(this.left2);
            this.groupBox1.Controls.Add(this.right2);
            this.groupBox1.Controls.Add(this.left1);
            this.groupBox1.Controls.Add(this.right1);
            this.groupBox1.Controls.Add(this.label23);
            this.groupBox1.Controls.Add(this.label21);
            this.groupBox1.Controls.Add(this.f2);
            this.groupBox1.Controls.Add(this.lens4);
            this.groupBox1.Controls.Add(this.label22);
            this.groupBox1.Controls.Add(this.f1);
            this.groupBox1.Controls.Add(this.label20);
            this.groupBox1.Controls.Add(this.lens3);
            this.groupBox1.Controls.Add(this.label19);
            this.groupBox1.Controls.Add(this.lens2);
            this.groupBox1.Controls.Add(this.label18);
            this.groupBox1.Controls.Add(this.lens1);
            this.groupBox1.Controls.Add(this.label17);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Location = new Point(0x15, 0x41);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new Size(0x19e, 0x12d);
            this.groupBox1.TabIndex = 0x13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Eye Details";
            this.groupBox1.Enter += new EventHandler(this.groupBox1_Enter);
            this.button9.Location = new Point(0x139, 0x103);
            this.button9.Name = "button9";
            this.button9.Size = new Size(0x29, 0x17);
            this.button9.TabIndex = 0x49;
            this.button9.Text = "Clear";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new EventHandler(this.button9_Click);
            this.left_6.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.left_6.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.left_6.FormattingEnabled = true;
            object[] objArray3 = new object[] { "N-6", "N-8", "N-10", "N-12", "N-18", "N-24", "N-36" };
            this.left_6.Items.AddRange(objArray3);
            this.left_6.Location = new Point(0x159, 70);
            this.left_6.Name = "left_6";
            this.left_6.Size = new Size(0x30, 0x15);
            this.left_6.TabIndex = 0x13;
            this.right_6.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.right_6.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.right_6.FormattingEnabled = true;
            object[] objArray4 = new object[] { "N-6", "N-8", "N-10", "N-12", "N-18", "N-24", "N-36" };
            this.right_6.Items.AddRange(objArray4);
            this.right_6.Location = new Point(0x159, 0x2d);
            this.right_6.Name = "right_6";
            this.right_6.Size = new Size(0x30, 0x15);
            this.right_6.TabIndex = 13;
            this.left_5.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.left_5.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.left_5.FormattingEnabled = true;
            object[] objArray5 = new object[] { "+1.00", "+1.25", "+1.50", "+1.75", "+2.00", "+2.25", "+2.50", "+2.75", "+3.00" };
            this.left_5.Items.AddRange(objArray5);
            this.left_5.Location = new Point(0x124, 70);
            this.left_5.Name = "left_5";
            this.left_5.Size = new Size(0x30, 0x15);
            this.left_5.TabIndex = 0x12;
            this.right_5.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.right_5.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.right_5.FormattingEnabled = true;
            object[] objArray6 = new object[] { "+1.00", "+1.25", "+1.50", "+1.75", "+2.00", "+2.25", "+2.50", "+2.75", "+3.00" };
            this.right_5.Items.AddRange(objArray6);
            this.right_5.Location = new Point(0x124, 0x2e);
            this.right_5.Name = "right_5";
            this.right_5.Size = new Size(0x30, 0x15);
            this.right_5.TabIndex = 12;
            this.left4.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.left4.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.left4.FormattingEnabled = true;
            object[] objArray7 = new object[] { "6/6", "6/9", "6/12", "6/18", "6/24", "6/36" };
            this.left4.Items.AddRange(objArray7);
            this.left4.Location = new Point(0xf1, 70);
            this.left4.Name = "left4";
            this.left4.Size = new Size(0x30, 0x15);
            this.left4.TabIndex = 0x11;
            this.right4.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.right4.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.right4.FormattingEnabled = true;
            object[] objArray8 = new object[] { "6/6", "6/9", "6/12", "6/18", "6/24", "6/36" };
            this.right4.Items.AddRange(objArray8);
            this.right4.Location = new Point(0xf1, 0x2e);
            this.right4.Name = "right4";
            this.right4.Size = new Size(0x30, 0x15);
            this.right4.TabIndex = 11;
            this.button8.Location = new Point(0x167, 0x103);
            this.button8.Name = "button8";
            this.button8.Size = new Size(0x2e, 0x17);
            this.button8.TabIndex = 0x35;
            this.button8.Text = "Delete";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new EventHandler(this.button8_Click);
            this.pen_status.AutoSize = true;
            this.pen_status.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.pen_status.ForeColor = Color.Lime;
            this.pen_status.Location = new Point(0x5f, 0x11d);
            this.pen_status.Name = "pen_status";
            this.pen_status.Size = new Size(0x13, 13);
            this.pen_status.TabIndex = 0x48;
            this.pen_status.Text = "---";
            this.pen_status.TextChanged += new EventHandler(this.pen_status_TextChanged);
            this.button5.Location = new Point(0xe8, 0x103);
            this.button5.Name = "button5";
            this.button5.Size = new Size(0x4b, 0x17);
            this.button5.TabIndex = 0x1f;
            this.button5.Text = "Update";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new EventHandler(this.button5_Click_1);
            this.p4.Enabled = false;
            this.p4.Location = new Point(0x128, 0xdd);
            this.p4.Name = "p4";
            this.p4.Size = new Size(100, 20);
            this.p4.TabIndex = 0x1d;
            this.p4.Text = "0";
            this.p4.TextChanged += new EventHandler(this.p4_TextChanged);
            this.p1.Location = new Point(80, 0xb2);
            this.p1.Multiline = true;
            this.p1.Name = "p1";
            this.p1.Size = new Size(0x76, 0x23);
            this.p1.TabIndex = 0x1a;
            this.label24.AutoSize = true;
            this.label24.Location = new Point(5, 0xb5);
            this.label24.Name = "label24";
            this.label24.Size = new Size(0x2c, 13);
            this.label24.TabIndex = 0x44;
            this.label24.Text = "Remark";
            this.button4.Location = new Point(0x97, 0x103);
            this.button4.Name = "button4";
            this.button4.Size = new Size(0x4b, 0x17);
            this.button4.TabIndex = 30;
            this.button4.Text = "Insert";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new EventHandler(this.button4_Click_1);
            this.label25.AutoSize = true;
            this.label25.Location = new Point(0xde, 0xb5);
            this.label25.Name = "label25";
            this.label25.Size = new Size(0x49, 13);
            this.label25.TabIndex = 0x3b;
            this.label25.Text = "Extra Charges";
            this.p5.AutoSize = true;
            this.p5.Location = new Point(0x5f, 0x108);
            this.p5.Name = "p5";
            this.p5.Size = new Size(0x10, 13);
            this.p5.TabIndex = 0x42;
            this.p5.Text = "---";
            this.p5.TextChanged += new EventHandler(this.p5_TextChanged);
            this.p2.Location = new Point(0x128, 0xb2);
            this.p2.Name = "p2";
            this.p2.Size = new Size(100, 20);
            this.p2.TabIndex = 0x1b;
            this.p2.Text = "0";
            this.p2.TextChanged += new EventHandler(this.p2_TextChanged);
            this.label26.AutoSize = true;
            this.label26.Location = new Point(6, 0xe0);
            this.label26.Name = "label26";
            this.label26.Size = new Size(0x43, 13);
            this.label26.TabIndex = 0x3d;
            this.label26.Text = "Paid Amount";
            this.label28.AutoSize = true;
            this.label28.Location = new Point(0xe0, 0xe0);
            this.label28.Name = "label28";
            this.label28.Size = new Size(70, 13);
            this.label28.TabIndex = 0x40;
            this.label28.Text = "Total Amount";
            this.p3.Location = new Point(80, 0xdd);
            this.p3.Name = "p3";
            this.p3.Size = new Size(100, 20);
            this.p3.TabIndex = 0x1c;
            this.p3.Text = "0";
            this.p3.TextChanged += new EventHandler(this.p3_TextChanged_1);
            this.label27.AutoSize = true;
            this.label27.Location = new Point(6, 0x108);
            this.label27.Name = "label27";
            this.label27.Size = new Size(0x55, 13);
            this.label27.TabIndex = 0x3f;
            this.label27.Text = "Pending Amount";
            this.left3.AccessibleName = "";
            this.left3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.left3.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.left3.FormattingEnabled = true;
            object[] objArray9 = new object[] { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            objArray9[9] = "10";
            objArray9[10] = "11";
            objArray9[11] = "12";
            objArray9[12] = "13";
            objArray9[13] = "14";
            objArray9[14] = "15";
            objArray9[15] = "16";
            objArray9[0x10] = "17";
            objArray9[0x11] = "18";
            objArray9[0x12] = "19";
            objArray9[0x13] = "20";
            objArray9[20] = "21";
            objArray9[0x15] = "22";
            objArray9[0x16] = "23";
            objArray9[0x17] = "24";
            objArray9[0x18] = "25";
            objArray9[0x19] = "26";
            objArray9[0x1a] = "27";
            objArray9[0x1b] = "28";
            objArray9[0x1c] = "29";
            objArray9[0x1d] = "30";
            objArray9[30] = "31";
            objArray9[0x1f] = "32";
            objArray9[0x20] = "33";
            objArray9[0x21] = "34";
            objArray9[0x22] = "35";
            objArray9[0x23] = "36";
            objArray9[0x24] = "37";
            objArray9[0x25] = "38";
            objArray9[0x26] = "39";
            objArray9[0x27] = "40";
            objArray9[40] = "41";
            objArray9[0x29] = "42";
            objArray9[0x2a] = "43";
            objArray9[0x2b] = "44";
            objArray9[0x2c] = "45";
            objArray9[0x2d] = "46";
            objArray9[0x2e] = "47";
            objArray9[0x2f] = "48";
            objArray9[0x30] = "49";
            objArray9[0x31] = "50";
            objArray9[50] = "51";
            objArray9[0x33] = "52";
            objArray9[0x34] = "53";
            objArray9[0x35] = "54";
            objArray9[0x36] = "55";
            objArray9[0x37] = "56";
            objArray9[0x38] = "57";
            objArray9[0x39] = "58";
            objArray9[0x3a] = "59";
            objArray9[0x3b] = "60";
            objArray9[60] = "61";
            objArray9[0x3d] = "62";
            objArray9[0x3e] = "63";
            objArray9[0x3f] = "64";
            objArray9[0x40] = "65";
            objArray9[0x41] = "66";
            objArray9[0x42] = "67";
            objArray9[0x43] = "68";
            objArray9[0x44] = "69";
            objArray9[0x45] = "70";
            objArray9[70] = "71";
            objArray9[0x47] = "72";
            objArray9[0x48] = "73";
            objArray9[0x49] = "74";
            objArray9[0x4a] = "75";
            objArray9[0x4b] = "76";
            objArray9[0x4c] = "77";
            objArray9[0x4d] = "78";
            objArray9[0x4e] = "79";
            objArray9[0x4f] = "80";
            objArray9[80] = "81";
            objArray9[0x51] = "82";
            objArray9[0x52] = "83";
            objArray9[0x53] = "84";
            objArray9[0x54] = "85";
            objArray9[0x55] = "86";
            objArray9[0x56] = "87";
            objArray9[0x57] = "88";
            objArray9[0x58] = "89";
            objArray9[0x59] = "90";
            objArray9[90] = "91";
            objArray9[0x5b] = "92";
            objArray9[0x5c] = "93";
            objArray9[0x5d] = "94";
            objArray9[0x5e] = "95";
            objArray9[0x5f] = "96";
            objArray9[0x60] = "97";
            objArray9[0x61] = "98";
            objArray9[0x62] = "99";
            objArray9[0x63] = "100";
            objArray9[100] = "101";
            objArray9[0x65] = "102";
            objArray9[0x66] = "103";
            objArray9[0x67] = "104";
            objArray9[0x68] = "105";
            objArray9[0x69] = "106";
            objArray9[0x6a] = "107";
            objArray9[0x6b] = "108";
            objArray9[0x6c] = "109";
            objArray9[0x6d] = "110";
            objArray9[110] = "111";
            objArray9[0x6f] = "112";
            objArray9[0x70] = "113";
            objArray9[0x71] = "114";
            objArray9[0x72] = "115";
            objArray9[0x73] = "116";
            objArray9[0x74] = "117";
            objArray9[0x75] = "118";
            objArray9[0x76] = "119";
            objArray9[0x77] = "120";
            objArray9[120] = "121";
            objArray9[0x79] = "122";
            objArray9[0x7a] = "123";
            objArray9[0x7b] = "124";
            objArray9[0x7c] = "125";
            objArray9[0x7d] = "126";
            objArray9[0x7e] = "127";
            objArray9[0x7f] = "128";
            objArray9[0x80] = "129";
            objArray9[0x81] = "130";
            objArray9[130] = "132";
            objArray9[0x83] = "132";
            objArray9[0x84] = "133";
            objArray9[0x85] = "134";
            objArray9[0x86] = "135";
            objArray9[0x87] = "136";
            objArray9[0x88] = "137";
            objArray9[0x89] = "138";
            objArray9[0x8a] = "139";
            objArray9[0x8b] = "140";
            objArray9[140] = "141";
            objArray9[0x8d] = "142";
            objArray9[0x8e] = "143";
            objArray9[0x8f] = "144";
            objArray9[0x90] = "145";
            objArray9[0x91] = "146";
            objArray9[0x92] = "147";
            objArray9[0x93] = "148";
            objArray9[0x94] = "149";
            objArray9[0x95] = "150";
            objArray9[150] = "151";
            objArray9[0x97] = "152";
            objArray9[0x98] = "153";
            objArray9[0x99] = "153";
            objArray9[0x9a] = "155";
            objArray9[0x9b] = "156";
            objArray9[0x9c] = "156";
            objArray9[0x9d] = "158";
            objArray9[0x9e] = "159";
            objArray9[0x9f] = "160";
            objArray9[160] = "161";
            objArray9[0xa1] = "162";
            objArray9[0xa2] = "163";
            objArray9[0xa3] = "164";
            objArray9[0xa4] = "165";
            objArray9[0xa5] = "166";
            objArray9[0xa6] = "167";
            objArray9[0xa7] = "168";
            objArray9[0xa8] = "179";
            objArray9[0xa9] = "170";
            objArray9[170] = "171";
            objArray9[0xab] = "172";
            objArray9[0xac] = "173";
            objArray9[0xad] = "174";
            objArray9[0xae] = "175";
            objArray9[0xaf] = "176";
            objArray9[0xb0] = "177";
            objArray9[0xb1] = "178";
            objArray9[0xb2] = "179";
            objArray9[0xb3] = "180";
            this.left3.Items.AddRange(objArray9);
            this.left3.Location = new Point(0xbb, 0x45);
            this.left3.Name = "left3";
            this.left3.Size = new Size(0x30, 0x15);
            this.left3.TabIndex = 0x10;
            this.right3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.right3.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.right3.FormattingEnabled = true;
            object[] objArray10 = new object[] { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            objArray10[9] = "10";
            objArray10[10] = "11";
            objArray10[11] = "12";
            objArray10[12] = "13";
            objArray10[13] = "14";
            objArray10[14] = "15";
            objArray10[15] = "16";
            objArray10[0x10] = "17";
            objArray10[0x11] = "18";
            objArray10[0x12] = "19";
            objArray10[0x13] = "20";
            objArray10[20] = "21";
            objArray10[0x15] = "22";
            objArray10[0x16] = "23";
            objArray10[0x17] = "24";
            objArray10[0x18] = "25";
            objArray10[0x19] = "26";
            objArray10[0x1a] = "27";
            objArray10[0x1b] = "28";
            objArray10[0x1c] = "29";
            objArray10[0x1d] = "30";
            objArray10[30] = "31";
            objArray10[0x1f] = "32";
            objArray10[0x20] = "33";
            objArray10[0x21] = "34";
            objArray10[0x22] = "35";
            objArray10[0x23] = "36";
            objArray10[0x24] = "37";
            objArray10[0x25] = "38";
            objArray10[0x26] = "39";
            objArray10[0x27] = "40";
            objArray10[40] = "41";
            objArray10[0x29] = "42";
            objArray10[0x2a] = "43";
            objArray10[0x2b] = "44";
            objArray10[0x2c] = "45";
            objArray10[0x2d] = "46";
            objArray10[0x2e] = "47";
            objArray10[0x2f] = "48";
            objArray10[0x30] = "49";
            objArray10[0x31] = "50";
            objArray10[50] = "51";
            objArray10[0x33] = "52";
            objArray10[0x34] = "53";
            objArray10[0x35] = "54";
            objArray10[0x36] = "55";
            objArray10[0x37] = "56";
            objArray10[0x38] = "57";
            objArray10[0x39] = "58";
            objArray10[0x3a] = "59";
            objArray10[0x3b] = "60";
            objArray10[60] = "61";
            objArray10[0x3d] = "62";
            objArray10[0x3e] = "63";
            objArray10[0x3f] = "64";
            objArray10[0x40] = "65";
            objArray10[0x41] = "66";
            objArray10[0x42] = "67";
            objArray10[0x43] = "68";
            objArray10[0x44] = "69";
            objArray10[0x45] = "70";
            objArray10[70] = "71";
            objArray10[0x47] = "72";
            objArray10[0x48] = "73";
            objArray10[0x49] = "74";
            objArray10[0x4a] = "75";
            objArray10[0x4b] = "76";
            objArray10[0x4c] = "77";
            objArray10[0x4d] = "78";
            objArray10[0x4e] = "79";
            objArray10[0x4f] = "80";
            objArray10[80] = "81";
            objArray10[0x51] = "82";
            objArray10[0x52] = "83";
            objArray10[0x53] = "84";
            objArray10[0x54] = "85";
            objArray10[0x55] = "86";
            objArray10[0x56] = "87";
            objArray10[0x57] = "88";
            objArray10[0x58] = "89";
            objArray10[0x59] = "90";
            objArray10[90] = "91";
            objArray10[0x5b] = "92";
            objArray10[0x5c] = "93";
            objArray10[0x5d] = "94";
            objArray10[0x5e] = "95";
            objArray10[0x5f] = "96";
            objArray10[0x60] = "97";
            objArray10[0x61] = "98";
            objArray10[0x62] = "99";
            objArray10[0x63] = "100";
            objArray10[100] = "101";
            objArray10[0x65] = "102";
            objArray10[0x66] = "103";
            objArray10[0x67] = "104";
            objArray10[0x68] = "105";
            objArray10[0x69] = "106";
            objArray10[0x6a] = "107";
            objArray10[0x6b] = "108";
            objArray10[0x6c] = "109";
            objArray10[0x6d] = "110";
            objArray10[110] = "111";
            objArray10[0x6f] = "112";
            objArray10[0x70] = "113";
            objArray10[0x71] = "114";
            objArray10[0x72] = "115";
            objArray10[0x73] = "116";
            objArray10[0x74] = "117";
            objArray10[0x75] = "118";
            objArray10[0x76] = "119";
            objArray10[0x77] = "120";
            objArray10[120] = "121";
            objArray10[0x79] = "122";
            objArray10[0x7a] = "123";
            objArray10[0x7b] = "124";
            objArray10[0x7c] = "125";
            objArray10[0x7d] = "126";
            objArray10[0x7e] = "127";
            objArray10[0x7f] = "128";
            objArray10[0x80] = "129";
            objArray10[0x81] = "130";
            objArray10[130] = "132";
            objArray10[0x83] = "132";
            objArray10[0x84] = "133";
            objArray10[0x85] = "134";
            objArray10[0x86] = "135";
            objArray10[0x87] = "136";
            objArray10[0x88] = "137";
            objArray10[0x89] = "138";
            objArray10[0x8a] = "139";
            objArray10[0x8b] = "140";
            objArray10[140] = "141";
            objArray10[0x8d] = "142";
            objArray10[0x8e] = "143";
            objArray10[0x8f] = "144";
            objArray10[0x90] = "145";
            objArray10[0x91] = "146";
            objArray10[0x92] = "147";
            objArray10[0x93] = "148";
            objArray10[0x94] = "149";
            objArray10[0x95] = "150";
            objArray10[150] = "151";
            objArray10[0x97] = "152";
            objArray10[0x98] = "153";
            objArray10[0x99] = "153";
            objArray10[0x9a] = "155";
            objArray10[0x9b] = "156";
            objArray10[0x9c] = "156";
            objArray10[0x9d] = "158";
            objArray10[0x9e] = "159";
            objArray10[0x9f] = "160";
            objArray10[160] = "161";
            objArray10[0xa1] = "162";
            objArray10[0xa2] = "163";
            objArray10[0xa3] = "164";
            objArray10[0xa4] = "165";
            objArray10[0xa5] = "166";
            objArray10[0xa6] = "167";
            objArray10[0xa7] = "168";
            objArray10[0xa8] = "179";
            objArray10[0xa9] = "170";
            objArray10[170] = "171";
            objArray10[0xab] = "172";
            objArray10[0xac] = "173";
            objArray10[0xad] = "174";
            objArray10[0xae] = "175";
            objArray10[0xaf] = "176";
            objArray10[0xb0] = "177";
            objArray10[0xb1] = "178";
            objArray10[0xb2] = "179";
            objArray10[0xb3] = "180";
            this.right3.Items.AddRange(objArray10);
            this.right3.Location = new Point(0xbb, 0x2d);
            this.right3.Name = "right3";
            this.right3.Size = new Size(0x30, 0x15);
            this.right3.TabIndex = 10;
            this.left2.AccessibleName = "";
            this.left2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.left2.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.left2.FormattingEnabled = true;
            object[] objArray11 = new object[] { "-8.00", "-7.75", "-7.50", "-7.25", "-7.00", "-6.75", "-6.50", "-6.25", "-6.00" };
            objArray11[9] = "-5.75";
            objArray11[10] = "-5.50";
            objArray11[11] = "-5.25";
            objArray11[12] = "-5.00";
            objArray11[13] = "-4.75";
            objArray11[14] = "-4.50";
            objArray11[15] = "-4.25";
            objArray11[0x10] = "-4.00";
            objArray11[0x11] = "-3.75";
            objArray11[0x12] = "-3.50";
            objArray11[0x13] = "-3.25";
            objArray11[20] = "-3.00";
            objArray11[0x15] = "-2.75";
            objArray11[0x16] = "-2.50";
            objArray11[0x17] = "-2.25";
            objArray11[0x18] = "-2.00";
            objArray11[0x19] = "-1.75";
            objArray11[0x1a] = "-1.50";
            objArray11[0x1b] = "-1.25";
            objArray11[0x1c] = "-1.00";
            objArray11[0x1d] = "-0.75";
            objArray11[30] = "-0.50";
            objArray11[0x1f] = "-0.25";
            objArray11[0x20] = "0.00";
            objArray11[0x21] = "+0.25";
            objArray11[0x22] = "+0.50";
            objArray11[0x23] = "+0.75";
            objArray11[0x24] = "+1.00";
            objArray11[0x25] = "+1.25";
            objArray11[0x26] = "+1.50";
            objArray11[0x27] = "+1.75";
            objArray11[40] = "+2.00";
            objArray11[0x29] = "+2.25";
            objArray11[0x2a] = "+2.50";
            objArray11[0x2b] = "+2.75";
            objArray11[0x2c] = "+3.00";
            objArray11[0x2d] = "+3.25";
            objArray11[0x2e] = "+3.50";
            objArray11[0x2f] = "+3.75";
            objArray11[0x30] = "+4.00";
            objArray11[0x31] = "+4.25";
            objArray11[50] = "+4.50";
            objArray11[0x33] = "+4.75";
            objArray11[0x34] = "+5.00";
            objArray11[0x35] = "+5.25";
            objArray11[0x36] = "+5.50";
            objArray11[0x37] = "+5.75";
            objArray11[0x38] = "+6.00";
            objArray11[0x39] = "+6.25";
            objArray11[0x3a] = "+6.50";
            objArray11[0x3b] = "+6.75";
            objArray11[60] = "+7.00";
            objArray11[0x3d] = "+7.25";
            objArray11[0x3e] = "+7.50";
            objArray11[0x3f] = "+7.75";
            objArray11[0x40] = "+8.00";
            this.left2.Items.AddRange(objArray11);
            this.left2.Location = new Point(0x86, 0x45);
            this.left2.Name = "left2";
            this.left2.Size = new Size(0x30, 0x15);
            this.left2.TabIndex = 15;
            this.right2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.right2.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.right2.FormattingEnabled = true;
            object[] objArray12 = new object[] { "-8.00", "-7.75", "-7.50", "-7.25", "-7.00", "-6.75", "-6.50", "-6.25", "-6.00" };
            objArray12[9] = "-5.75";
            objArray12[10] = "-5.50";
            objArray12[11] = "-5.25";
            objArray12[12] = "-5.00";
            objArray12[13] = "-4.75";
            objArray12[14] = "-4.50";
            objArray12[15] = "-4.25";
            objArray12[0x10] = "-4.00";
            objArray12[0x11] = "-3.75";
            objArray12[0x12] = "-3.50";
            objArray12[0x13] = "-3.25";
            objArray12[20] = "-3.00";
            objArray12[0x15] = "-2.75";
            objArray12[0x16] = "-2.50";
            objArray12[0x17] = "-2.25";
            objArray12[0x18] = "-2.00";
            objArray12[0x19] = "-1.75";
            objArray12[0x1a] = "-1.50";
            objArray12[0x1b] = "-1.25";
            objArray12[0x1c] = "-1.00";
            objArray12[0x1d] = "-0.75";
            objArray12[30] = "-0.50";
            objArray12[0x1f] = "-0.25";
            objArray12[0x20] = "0.00";
            objArray12[0x21] = "+0.25";
            objArray12[0x22] = "+0.50";
            objArray12[0x23] = "+0.75";
            objArray12[0x24] = "+1.00";
            objArray12[0x25] = "+1.25";
            objArray12[0x26] = "+1.50";
            objArray12[0x27] = "+1.75";
            objArray12[40] = "+2.00";
            objArray12[0x29] = "+2.25";
            objArray12[0x2a] = "+2.50";
            objArray12[0x2b] = "+2.75";
            objArray12[0x2c] = "+3.00";
            objArray12[0x2d] = "+3.25";
            objArray12[0x2e] = "+3.50";
            objArray12[0x2f] = "+3.75";
            objArray12[0x30] = "+4.00";
            objArray12[0x31] = "+4.25";
            objArray12[50] = "+4.50";
            objArray12[0x33] = "+4.75";
            objArray12[0x34] = "+5.00";
            objArray12[0x35] = "+5.25";
            objArray12[0x36] = "+5.50";
            objArray12[0x37] = "+5.75";
            objArray12[0x38] = "+6.00";
            objArray12[0x39] = "+6.25";
            objArray12[0x3a] = "+6.50";
            objArray12[0x3b] = "+6.75";
            objArray12[60] = "+7.00";
            objArray12[0x3d] = "+7.25";
            objArray12[0x3e] = "+7.50";
            objArray12[0x3f] = "+7.75";
            objArray12[0x40] = "+8.00";
            this.right2.Items.AddRange(objArray12);
            this.right2.Location = new Point(0x86, 0x2d);
            this.right2.Name = "right2";
            this.right2.Size = new Size(0x30, 0x15);
            this.right2.TabIndex = 9;
            this.left1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.left1.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.left1.BackColor = SystemColors.Window;
            this.left1.ForeColor = SystemColors.MenuText;
            this.left1.FormattingEnabled = true;
            object[] objArray13 = new object[] { "-8.00", "-7.75", "-7.50", "-7.25", "-7.00", "-6.75", "-6.50", "-6.25", "-6.00" };
            objArray13[9] = "-5.75";
            objArray13[10] = "-5.50";
            objArray13[11] = "-5.25";
            objArray13[12] = "-5.00";
            objArray13[13] = "-4.75";
            objArray13[14] = "-4.50";
            objArray13[15] = "-4.25";
            objArray13[0x10] = "-4.00";
            objArray13[0x11] = "-3.75";
            objArray13[0x12] = "-3.50";
            objArray13[0x13] = "-3.25";
            objArray13[20] = "-3.00";
            objArray13[0x15] = "-2.75";
            objArray13[0x16] = "-2.50";
            objArray13[0x17] = "-2.25";
            objArray13[0x18] = "-2.00";
            objArray13[0x19] = "-1.75";
            objArray13[0x1a] = "-1.50";
            objArray13[0x1b] = "-1.25";
            objArray13[0x1c] = "-1.00";
            objArray13[0x1d] = "-0.75";
            objArray13[30] = "-0.50";
            objArray13[0x1f] = "-0.25";
            objArray13[0x20] = "0.00";
            objArray13[0x21] = "+0.25";
            objArray13[0x22] = "+0.50";
            objArray13[0x23] = "+0.75";
            objArray13[0x24] = "+1.00";
            objArray13[0x25] = "+1.25";
            objArray13[0x26] = "+1.50";
            objArray13[0x27] = "+1.75";
            objArray13[40] = "+2.00";
            objArray13[0x29] = "+2.25";
            objArray13[0x2a] = "+2.50";
            objArray13[0x2b] = "+2.75";
            objArray13[0x2c] = "+3.00";
            objArray13[0x2d] = "+3.25";
            objArray13[0x2e] = "+3.50";
            objArray13[0x2f] = "+3.75";
            objArray13[0x30] = "+4.00";
            objArray13[0x31] = "+4.25";
            objArray13[50] = "+4.50";
            objArray13[0x33] = "+4.75";
            objArray13[0x34] = "+5.00";
            objArray13[0x35] = "+5.25";
            objArray13[0x36] = "+5.50";
            objArray13[0x37] = "+5.75";
            objArray13[0x38] = "+6.00";
            objArray13[0x39] = "+6.25";
            objArray13[0x3a] = "+6.50";
            objArray13[0x3b] = "+6.75";
            objArray13[60] = "+7.00";
            objArray13[0x3d] = "+7.25";
            objArray13[0x3e] = "+7.50";
            objArray13[0x3f] = "+7.75";
            objArray13[0x40] = "+8.00";
            this.left1.Items.AddRange(objArray13);
            this.left1.Location = new Point(80, 70);
            this.left1.Name = "left1";
            this.left1.Size = new Size(0x30, 0x15);
            this.left1.TabIndex = 14;
            this.right1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.right1.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.right1.BackColor = SystemColors.Window;
            this.right1.ForeColor = SystemColors.MenuText;
            this.right1.FormattingEnabled = true;
            object[] objArray14 = new object[] { "-8.00", "-7.75", "-7.50", "-7.25", "-7.00", "-6.75", "-6.50", "-6.25", "-6.00" };
            objArray14[9] = "-5.75";
            objArray14[10] = "-5.50";
            objArray14[11] = "-5.25";
            objArray14[12] = "-5.00";
            objArray14[13] = "-4.75";
            objArray14[14] = "-4.50";
            objArray14[15] = "-4.25";
            objArray14[0x10] = "-4.00";
            objArray14[0x11] = "-3.75";
            objArray14[0x12] = "-3.50";
            objArray14[0x13] = "-3.25";
            objArray14[20] = "-3.00";
            objArray14[0x15] = "-2.75";
            objArray14[0x16] = "-2.50";
            objArray14[0x17] = "-2.25";
            objArray14[0x18] = "-2.00";
            objArray14[0x19] = "-1.75";
            objArray14[0x1a] = "-1.50";
            objArray14[0x1b] = "-1.25";
            objArray14[0x1c] = "-1.00";
            objArray14[0x1d] = "-0.75";
            objArray14[30] = "-0.50";
            objArray14[0x1f] = "-0.25";
            objArray14[0x20] = "0.00";
            objArray14[0x21] = "+0.25";
            objArray14[0x22] = "+0.50";
            objArray14[0x23] = "+0.75";
            objArray14[0x24] = "+1.00";
            objArray14[0x25] = "+1.25";
            objArray14[0x26] = "+1.50";
            objArray14[0x27] = "+1.75";
            objArray14[40] = "+2.00";
            objArray14[0x29] = "+2.25";
            objArray14[0x2a] = "+2.50";
            objArray14[0x2b] = "+2.75";
            objArray14[0x2c] = "+3.00";
            objArray14[0x2d] = "+3.25";
            objArray14[0x2e] = "+3.50";
            objArray14[0x2f] = "+3.75";
            objArray14[0x30] = "+4.00";
            objArray14[0x31] = "+4.25";
            objArray14[50] = "+4.50";
            objArray14[0x33] = "+4.75";
            objArray14[0x34] = "+5.00";
            objArray14[0x35] = "+5.25";
            objArray14[0x36] = "+5.50";
            objArray14[0x37] = "+5.75";
            objArray14[0x38] = "+6.00";
            objArray14[0x39] = "+6.25";
            objArray14[0x3a] = "+6.50";
            objArray14[0x3b] = "+6.75";
            objArray14[60] = "+7.00";
            objArray14[0x3d] = "+7.25";
            objArray14[0x3e] = "+7.50";
            objArray14[0x3f] = "+7.75";
            objArray14[0x40] = "+8.00";
            objArray14[0x41] = "+8.25";
            objArray14[0x42] = "+8.50";
            objArray14[0x43] = "+8.75";
            objArray14[0x44] = "+9.00";
            objArray14[0x45] = "+9.25";
            objArray14[70] = "+9.50";
            objArray14[0x47] = "+9.75";
            objArray14[0x48] = "+10.00";
            this.right1.Items.AddRange(objArray14);
            this.right1.Location = new Point(80, 0x2e);
            this.right1.Name = "right1";
            this.right1.Size = new Size(0x30, 0x15);
            this.right1.TabIndex = 8;
            this.label23.AutoSize = true;
            this.label23.Location = new Point(0xdf, 0x97);
            this.label23.Name = "label23";
            this.label23.Size = new Size(0x3f, 13);
            this.label23.TabIndex = 50;
            this.label23.Text = "Frame Price";
            this.label21.AutoSize = true;
            this.label21.Location = new Point(0xdf, 0x7d);
            this.label21.Name = "label21";
            this.label21.Size = new Size(0x39, 13);
            this.label21.TabIndex = 0x30;
            this.label21.Text = "Lens Price";
            this.f2.Location = new Point(0x128, 0x94);
            this.f2.Name = "f2";
            this.f2.Size = new Size(100, 20);
            this.f2.TabIndex = 0x19;
            this.f2.Text = "0";
            this.f2.TextChanged += new EventHandler(this.f2_TextChanged);
            this.lens4.Location = new Point(0x128, 0x7a);
            this.lens4.Name = "lens4";
            this.lens4.Size = new Size(100, 20);
            this.lens4.TabIndex = 0x17;
            this.lens4.Text = "0";
            this.lens4.TextChanged += new EventHandler(this.lens4_TextChanged);
            this.label22.AutoSize = true;
            this.label22.Location = new Point(6, 0x98);
            this.label22.Name = "label22";
            this.label22.Size = new Size(0x3f, 13);
            this.label22.TabIndex = 0x2c;
            this.label22.Text = "Frame Type";
            this.f1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.f1.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.f1.FormattingEnabled = true;
            object[] objArray15 = new object[] { "RIMLESS", "FULL-FRAME", "SUPRA", "POLY C.", "GOGGLE", "INDIAN", "SHEET" };
            this.f1.Items.AddRange(objArray15);
            this.f1.Location = new Point(80, 0x95);
            this.f1.Name = "f1";
            this.f1.Size = new Size(100, 0x15);
            this.f1.TabIndex = 0x18;
            this.lens1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.lens1.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.lens1.FormattingEnabled = true;
            object[] objArray16 = new object[] { "GLASS" };
            this.lens1.Items.AddRange(objArray16);
            this.lens1.Location = new Point(80, 0x60);
            this.lens1.Name = "lens1";
            this.lens1.Size = new Size(100, 0x15);
            this.lens1.TabIndex = 20;
            this.label14.AutoSize = true;
            this.label14.Location = new Point(0xb9, 0x1a);
            this.label14.Name = "label14";
            this.label14.Size = new Size(0x1f, 13);
            this.label14.TabIndex = 30;
            this.label14.Text = "AXIS";
            this.label13.AutoSize = true;
            this.label13.Location = new Point(0x83, 0x1a);
            this.label13.Name = "label13";
            this.label13.Size = new Size(30, 13);
            this.label13.TabIndex = 0x1b;
            this.label13.Text = "CYL.";
            this.label12.AutoSize = true;
            this.label12.Location = new Point(6, 0x49);
            this.label12.Name = "label12";
            this.label12.Size = new Size(0x2e, 13);
            this.label12.TabIndex = 0x19;
            this.label12.Text = "OS(Left)";
            this.label11.AutoSize = true;
            this.label11.Location = new Point(6, 0x31);
            this.label11.Name = "label11";
            this.label11.Size = new Size(0x36, 13);
            this.label11.TabIndex = 0x17;
            this.label11.Text = "OD(Right)";
            this.label10.AutoSize = true;
            this.label10.Location = new Point(0x4f, 0x1a);
            this.label10.Name = "label10";
            this.label10.Size = new Size(0x20, 13);
            this.label10.TabIndex = 0x16;
            this.label10.Text = "SPH.";
            this.label9.AutoSize = true;
            this.label9.Location = new Point(6, 0x1a);
            this.label9.Name = "label9";
            this.label9.Size = new Size(0x19, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "Eye";
            this.t7.Location = new Point(0x65, 0x87);
            this.t7.Name = "t7";
            this.t7.Size = new Size(100, 20);
            this.t7.TabIndex = 8;
            this.button1.Location = new Point(0x61, 0xf1);
            this.button1.Name = "button1";
            this.button1.Size = new Size(0x4b, 0x17);
            this.button1.TabIndex = 9;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new EventHandler(this.button1_Click);
            this.label8.AutoSize = true;
            this.label8.Location = new Point(0x12, 0x8a);
            this.label8.Name = "label8";
            this.label8.Size = new Size(0x1a, 13);
            this.label8.TabIndex = 0x11;
            this.label8.Text = "Age";
            this.t6.Location = new Point(0x65, 0x69);
            this.t6.Name = "t6";
            this.t6.Size = new Size(100, 20);
            this.t6.TabIndex = 7;
            this.t6.TextChanged += new EventHandler(this.t6_TextChanged);
            this.label7.AutoSize = true;
            this.label7.Location = new Point(0x12, 0x6c);
            this.label7.Name = "label7";
            this.label7.Size = new Size(0x3d, 13);
            this.label7.TabIndex = 15;
            this.label7.Text = "Contact No";
            this.r2.AutoSize = true;
            this.r2.Location = new Point(0x175, 0x6d);
            this.r2.Name = "r2";
            this.r2.Size = new Size(0x3b, 0x11);
            this.r2.TabIndex = 14;
            this.r2.Text = "Female";
            this.r2.UseVisualStyleBackColor = true;
            this.t5.Location = new Point(0x65, 0x4e);
            this.t5.Name = "t5";
            this.t5.Size = new Size(100, 20);
            this.t5.TabIndex = 6;
            this.t3.Location = new Point(0x65, 0x2e);
            this.t3.Name = "t3";
            this.t3.Size = new Size(100, 20);
            this.t3.TabIndex = 4;
            this.label3.AutoSize = true;
            this.label3.Location = new Point(0x12, 0x31);
            this.label3.Name = "label3";
            this.label3.Size = new Size(0x23, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Name";
            this.label2.AutoSize = true;
            this.label2.Location = new Point(0xea, 0x31);
            this.label2.Name = "label2";
            this.label2.Size = new Size(0x44, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Examined By";
            this.tabPage1.Controls.Add(this.button7);
            this.tabPage1.Controls.Add(this.label43);
            this.tabPage1.Controls.Add(this.label42);
            this.tabPage1.Controls.Add(this.label40);
            this.tabPage1.Controls.Add(this.label33);
            this.tabPage1.Controls.Add(this.label32);
            this.tabPage1.Controls.Add(this.d4);
            this.tabPage1.Controls.Add(this.button3);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.t7);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.t6);
            this.tabPage1.Controls.Add(this.label7);
            this.tabPage1.Controls.Add(this.r2);
            this.tabPage1.Controls.Add(this.r1);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.t5);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.t3);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.t2);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Location = new Point(4, 0x16);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new Padding(3);
            this.tabPage1.Size = new Size(0x1c4, 0x25c);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "(+) Add New Information";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.button7.Location = new Point(340, 0xf1);
            this.button7.Name = "button7";
            this.button7.Size = new Size(0x4b, 0x17);
            this.button7.TabIndex = 0x34;
            this.button7.Text = "Delete";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new EventHandler(this.button7_Click);
            this.label43.AutoSize = true;
            this.label43.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.label43.ForeColor = Color.Red;
            this.label43.Location = new Point(0x114, 0x6a);
            this.label43.Name = "label43";
            this.label43.Size = new Size(15, 20);
            this.label43.TabIndex = 0x2b;
            this.label43.Text = "*";
            this.label42.AutoSize = true;
            this.label42.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.label42.ForeColor = Color.Red;
            this.label42.Location = new Point(0x4f, 0x69);
            this.label42.Name = "label42";
            this.label42.Size = new Size(15, 20);
            this.label42.TabIndex = 0x2a;
            this.label42.Text = "*";
            this.label40.AutoSize = true;
            this.label40.Font = new Font("Microsoft Sans Serif", 12f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.label40.ForeColor = Color.Red;
            this.label40.Location = new Point(0x30, 0x2e);
            this.label40.Name = "label40";
            this.label40.Size = new Size(15, 20);
            this.label40.TabIndex = 0x29;
            this.label40.Text = "*";
            this.label33.AutoSize = true;
            this.label33.Font = new Font("Microsoft Sans Serif", 15.75f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.label33.ForeColor = Color.Red;
            this.label33.Location = new Point(0x10, 320);
            this.label33.Name = "label33";
            this.label33.Size = new Size(20, 0x19);
            this.label33.TabIndex = 0x26;
            this.label33.Text = "*";
            this.label32.AutoSize = true;
            this.label32.Location = new Point(6, 0x141);
            this.label32.Name = "label32";
            this.label32.Size = new Size(0xae, 13);
            this.label32.TabIndex = 0x25;
            this.label32.Text = "(        )Sign indicates required fields.";
            this.d4.CustomFormat = "";
            this.d4.Format = DateTimePickerFormat.Short;
            this.d4.Location = new Point(0x13d, 0x4a);
            this.d4.Name = "d4";
            this.d4.Size = new Size(0x73, 20);
            this.d4.TabIndex = 5;
            this.d4.Value = new DateTime(0x7e1, 8, 14, 0, 0, 0, 0);
            this.button3.Location = new Point(0x103, 0xf1);
            this.button3.Name = "button3";
            this.button3.Size = new Size(0x4b, 0x17);
            this.button3.TabIndex = 0x20;
            this.button3.Text = "Update";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new EventHandler(this.button3_Click);
            this.button2.Location = new Point(0xb2, 0xf1);
            this.button2.Name = "button2";
            this.button2.Size = new Size(0x4b, 0x17);
            this.button2.TabIndex = 0x1f;
            this.button2.Text = "Clear";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new EventHandler(this.button2_Click);
            this.r1.AutoSize = true;
            this.r1.Checked = true;
            this.r1.Location = new Point(0x13d, 0x6d);
            this.r1.Name = "r1";
            this.r1.Size = new Size(0x30, 0x11);
            this.r1.TabIndex = 13;
            this.r1.TabStop = true;
            this.r1.Text = "Male";
            this.r1.UseVisualStyleBackColor = true;
            this.label6.AutoSize = true;
            this.label6.Location = new Point(0xea, 0x6f);
            this.label6.Name = "label6";
            this.label6.Size = new Size(0x2a, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Gender";
            this.label5.AutoSize = true;
            this.label5.Location = new Point(0x12, 0x51);
            this.label5.Name = "label5";
            this.label5.Size = new Size(0x2d, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Address";
            this.label4.AutoSize = true;
            this.label4.Location = new Point(0xea, 0x51);
            this.label4.Name = "label4";
            this.label4.Size = new Size(0x36, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Birth Date";
            this.t2.Location = new Point(0x13d, 0x2e);
            this.t2.Name = "t2";
            this.t2.Size = new Size(100, 20);
            this.t2.TabIndex = 3;
            this.cust_id.AutoSize = true;
            this.cust_id.Location = new Point(0x6b, 0x18);
            this.cust_id.Name = "cust_id";
            this.cust_id.Size = new Size(0x10, 13);
            this.cust_id.TabIndex = 40;
            this.cust_id.Text = "---";
            this.label29.AutoSize = true;
            this.label29.Location = new Point(0x1a, 0x18);
            this.label29.Name = "label29";
            this.label29.Size = new Size(0x44, 13);
            this.label29.TabIndex = 0x23;
            this.label29.Text = "Customer ID.";
            this.tabPage2.Controls.Add(this.label46);
            this.tabPage2.Controls.Add(this.label45);
            this.tabPage2.Controls.Add(this.Pre_id);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Controls.Add(this.cust_id);
            this.tabPage2.Controls.Add(this.label29);
            this.tabPage2.Location = new Point(4, 0x16);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new Padding(3);
            this.tabPage2.Size = new Size(0x1c4, 0x25c);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Add New Prescription";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.label46.AutoSize = true;
            this.label46.Font = new Font("Microsoft Sans Serif", 6.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.label46.Location = new Point(0x10f, 0x2c);
            this.label46.Name = "label46";
            this.label46.Size = new Size(140, 12);
            this.label46.TabIndex = 0x48;
            this.label46.Text = "*Only Use for Updation of Petient";
            this.label45.AutoSize = true;
            this.label45.Location = new Point(0xe4, 0x18);
            this.label45.Name = "label45";
            this.label45.Size = new Size(0x4c, 13);
            this.label45.TabIndex = 70;
            this.label45.Text = "Prescription ID";
            this.Pre_id.Location = new Point(0x137, 0x15);
            this.Pre_id.Name = "Pre_id";
            this.Pre_id.Size = new Size(100, 20);
            this.Pre_id.TabIndex = 0x47;
            this.groupBox2.Controls.Add(this.dataGridView2);
            this.groupBox2.Font = new Font("Microsoft Sans Serif", 9.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.groupBox2.Location = new Point(0x15, 0x175);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new Size(0x19e, 0xbd);
            this.groupBox2.TabIndex = 0x2a;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Prescription History";
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new Point(12, 0x19);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new Size(0x18c, 0x9e);
            this.dataGridView2.TabIndex = 0x29;
            this.dataGridView2.CellMouseClick += new DataGridViewCellMouseEventHandler(this.dataGridView2_CellMouseClick);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new Point(12, 0x31);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new Size(460, 630);
            this.tabControl1.TabIndex = 7;
            ToolStripItem[] toolStripItems = new ToolStripItem[] { this.fileToolStripMenuItem, this.editToolStripMenuItem, this.toolToolStripMenuItem, this.helpToolStripMenuItem };
            this.menuStrip1.Items.AddRange(toolStripItems);
            this.menuStrip1.Location = new Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new Size(0x4f7, 0x18);
            this.menuStrip1.TabIndex = 8;
            this.menuStrip1.Text = "menuStrip1";
            ToolStripItem[] itemArray2 = new ToolStripItem[] { this.saveAsToolStripMenuItem, this.sendToToolStripMenuItem, this.exitToolStripMenuItem, this.exitToolStripMenuItem1 };
            this.fileToolStripMenuItem.DropDownItems.AddRange(itemArray2);
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new Size(0x25, 20);
            this.fileToolStripMenuItem.Text = "File";
            ToolStripItem[] itemArray3 = new ToolStripItem[] { this.excelSheetToolStripMenuItem, this.accessFIleToolStripMenuItem };
            this.saveAsToolStripMenuItem.DropDownItems.AddRange(itemArray3);
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new Size(0x7b, 0x16);
            this.saveAsToolStripMenuItem.Text = "Save As...";
            this.excelSheetToolStripMenuItem.Name = "excelSheetToolStripMenuItem";
            this.excelSheetToolStripMenuItem.Size = new Size(0x84, 0x16);
            this.excelSheetToolStripMenuItem.Text = "Excel Sheet";
            this.accessFIleToolStripMenuItem.Name = "accessFIleToolStripMenuItem";
            this.accessFIleToolStripMenuItem.Size = new Size(0x84, 0x16);
            this.accessFIleToolStripMenuItem.Text = "Access FIle";
            this.sendToToolStripMenuItem.Name = "sendToToolStripMenuItem";
            this.sendToToolStripMenuItem.Size = new Size(0x7b, 0x16);
            this.sendToToolStripMenuItem.Text = "Send To";
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new Size(0x7b, 0x16);
            this.exitToolStripMenuItem.Text = "Print";
            this.exitToolStripMenuItem1.Name = "exitToolStripMenuItem1";
            this.exitToolStripMenuItem1.Size = new Size(0x7b, 0x16);
            this.exitToolStripMenuItem1.Text = "Exit.";
            this.exitToolStripMenuItem1.Click += new EventHandler(this.exitToolStripMenuItem1_Click);
            this.editToolStripMenuItem.Name = "editToolStripMenuItem";
            this.editToolStripMenuItem.Size = new Size(0x27, 20);
            this.editToolStripMenuItem.Text = "Edit";
            ToolStripItem[] itemArray4 = new ToolStripItem[] { this.backupToolStripMenuItem };
            this.toolToolStripMenuItem.DropDownItems.AddRange(itemArray4);
            this.toolToolStripMenuItem.Name = "toolToolStripMenuItem";
            this.toolToolStripMenuItem.Size = new Size(0x2a, 20);
            this.toolToolStripMenuItem.Text = "Tool";
            ToolStripItem[] itemArray5 = new ToolStripItem[] { this.exportDatabaseToolStripMenuItem, this.importDatabaseToolStripMenuItem };
            this.backupToolStripMenuItem.DropDownItems.AddRange(itemArray5);
            this.backupToolStripMenuItem.Name = "backupToolStripMenuItem";
            this.backupToolStripMenuItem.Size = new Size(0x71, 0x16);
            this.backupToolStripMenuItem.Text = "Backup";
            this.exportDatabaseToolStripMenuItem.Name = "exportDatabaseToolStripMenuItem";
            this.exportDatabaseToolStripMenuItem.Size = new Size(0xa1, 0x16);
            this.exportDatabaseToolStripMenuItem.Text = "Export Database";
            this.exportDatabaseToolStripMenuItem.Click += new EventHandler(this.exportDatabaseToolStripMenuItem_Click);
            this.importDatabaseToolStripMenuItem.Name = "importDatabaseToolStripMenuItem";
            this.importDatabaseToolStripMenuItem.Size = new Size(0xa1, 0x16);
            this.importDatabaseToolStripMenuItem.Text = "Import Database";
            this.importDatabaseToolStripMenuItem.Click += new EventHandler(this.importDatabaseToolStripMenuItem_Click);
            ToolStripItem[] itemArray6 = new ToolStripItem[] { this.helpToolStripMenuItem1, this.aboutToolStripMenuItem };
            this.helpToolStripMenuItem.DropDownItems.AddRange(itemArray6);
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new Size(0x2c, 20);
            this.helpToolStripMenuItem.Text = "Help";
            this.helpToolStripMenuItem1.Name = "helpToolStripMenuItem1";
            this.helpToolStripMenuItem1.Size = new Size(0x6b, 0x16);
            this.helpToolStripMenuItem1.Text = "Help";
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new Size(0x6b, 0x16);
            this.aboutToolStripMenuItem.Text = "About";
            this.groupBox3.Controls.Add(this.button6);
            this.groupBox3.Controls.Add(this.Search_by);
            this.groupBox3.Controls.Add(this.Search_txt);
            this.groupBox3.Controls.Add(this.dataGridView1);
            this.groupBox3.Controls.Add(this.label44);
            this.groupBox3.Font = new Font("Microsoft Sans Serif", 9.75f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.groupBox3.Location = new Point(0x1de, 0x31);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new Size(0x315, 0x177);
            this.groupBox3.TabIndex = 9;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Patient - History";
            this.button6.Location = new Point(0x2b7, 0x13);
            this.button6.Name = "button6";
            this.button6.Size = new Size(0x4b, 0x17);
            this.button6.TabIndex = 0x2c;
            this.button6.Text = "Refresh";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new EventHandler(this.button6_Click);
            this.Search_by.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.Search_by.AutoCompleteSource = AutoCompleteSource.ListItems;
            this.Search_by.FormattingEnabled = true;
            object[] objArray17 = new object[] { "Name", "Pending Bill", "Examined By", "Contact no", "Order date" };
            this.Search_by.Items.AddRange(objArray17);
            this.Search_by.Location = new Point(0xde, 0x13);
            this.Search_by.Name = "Search_by";
            this.Search_by.Size = new Size(100, 0x18);
            this.Search_by.TabIndex = 0x33;
            this.Search_by.Text = "Name";
            this.Search_txt.Location = new Point(0x73, 0x13);
            this.Search_txt.Name = "Search_txt";
            this.Search_txt.Size = new Size(100, 0x16);
            this.Search_txt.TabIndex = 0x20;
            this.Search_txt.TextChanged += new EventHandler(this.Search_txt_TextChanged);
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new Point(12, 0x2e);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new Size(0x2f6, 0x143);
            this.dataGridView1.TabIndex = 10;
            this.dataGridView1.CellContentClick += new DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellContentDoubleClick += new DataGridViewCellEventHandler(this.dataGridView1_CellContentDoubleClick);
            this.dataGridView1.CellMouseClick += new DataGridViewCellMouseEventHandler(this.dataGridView1_CellMouseClick);
            this.label44.AutoSize = true;
            this.label44.Location = new Point(20, 0x16);
            this.label44.Name = "label44";
            this.label44.Size = new Size(0x62, 0x10);
            this.label44.TabIndex = 0x2d;
            this.label44.Text = "Search Patient:";
            this.groupBox4.Controls.Add(this.label47);
            this.groupBox4.Controls.Add(this.pre_id2);
            this.groupBox4.Controls.Add(this.lb_id);
            this.groupBox4.Controls.Add(this.label41);
            this.groupBox4.Controls.Add(this.groupBox6);
            this.groupBox4.Controls.Add(this.groupBox5);
            this.groupBox4.Controls.Add(this.lb_contact);
            this.groupBox4.Controls.Add(this.lb_name);
            this.groupBox4.Controls.Add(this.label30);
            this.groupBox4.Controls.Add(this.label31);
            this.groupBox4.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.groupBox4.Location = new Point(0x1de, 0x1ac);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new Size(0x30d, 0xf7);
            this.groupBox4.TabIndex = 10;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Lens Details of Customer";
            this.label47.AutoSize = true;
            this.label47.Location = new Point(0x278, 0x11);
            this.label47.Name = "label47";
            this.label47.Size = new Size(0x57, 15);
            this.label47.TabIndex = 0x55;
            this.label47.Text = "Prescription ID";
            this.pre_id2.Location = new Point(0x2d5, 14);
            this.pre_id2.Name = "pre_id2";
            this.pre_id2.Size = new Size(0x2c, 0x15);
            this.pre_id2.TabIndex = 0x56;
            this.lb_id.AutoSize = true;
            this.lb_id.Location = new Point(100, 0x11);
            this.lb_id.Name = "lb_id";
            this.lb_id.Size = new Size(0x13, 15);
            this.lb_id.TabIndex = 0x27;
            this.lb_id.Text = "---";
            this.label41.AutoSize = true;
            this.label41.Location = new Point(0x11, 0x11);
            this.label41.Name = "label41";
            this.label41.Size = new Size(0x4b, 15);
            this.label41.TabIndex = 0x26;
            this.label41.Text = "Customer ID";
            this.groupBox6.BackColor = SystemColors.ButtonHighlight;
            this.groupBox6.Controls.Add(this.button10);
            this.groupBox6.Controls.Add(this.label1);
            this.groupBox6.Controls.Add(this.label53);
            this.groupBox6.Controls.Add(this.total2);
            this.groupBox6.Controls.Add(this.label52);
            this.groupBox6.Controls.Add(this.total1);
            this.groupBox6.Controls.Add(this.label51);
            this.groupBox6.Controls.Add(this.total3);
            this.groupBox6.Controls.Add(this.ll6);
            this.groupBox6.Controls.Add(this.lr6);
            this.groupBox6.Controls.Add(this.ll5);
            this.groupBox6.Controls.Add(this.lr5);
            this.groupBox6.Controls.Add(this.ll4);
            this.groupBox6.Controls.Add(this.lr4);
            this.groupBox6.Controls.Add(this.ll3);
            this.groupBox6.Controls.Add(this.lr3);
            this.groupBox6.Controls.Add(this.ll2);
            this.groupBox6.Controls.Add(this.lr2);
            this.groupBox6.Controls.Add(this.ll1);
            this.groupBox6.Controls.Add(this.lr1);
            this.groupBox6.Controls.Add(this.fm_type);
            this.groupBox6.Controls.Add(this.lb_side);
            this.groupBox6.Controls.Add(this.lb_for);
            this.groupBox6.Controls.Add(this.lb_type);
            this.groupBox6.Controls.Add(this.label58);
            this.groupBox6.Controls.Add(this.label59);
            this.groupBox6.Controls.Add(this.label60);
            this.groupBox6.Controls.Add(this.label62);
            this.groupBox6.Controls.Add(this.label49);
            this.groupBox6.Controls.Add(this.label48);
            this.groupBox6.Controls.Add(this.label50);
            this.groupBox6.Controls.Add(this.label34);
            this.groupBox6.Controls.Add(this.label39);
            this.groupBox6.Controls.Add(this.label35);
            this.groupBox6.Controls.Add(this.label38);
            this.groupBox6.Controls.Add(this.label36);
            this.groupBox6.Controls.Add(this.label37);
            this.groupBox6.Location = new Point(12, 0x31);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new Size(0x2ed, 0xc0);
            this.groupBox6.TabIndex = 0x24;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Priscription";
            this.button10.BackColor = Color.AliceBlue;
            this.button10.Enabled = false;
            this.button10.FlatStyle = FlatStyle.Popup;
            this.button10.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.button10.Location = new Point(0x188, 0xa2);
            this.button10.Name = "button10";
            this.button10.Size = new Size(0x3e, 0x17);
            this.button10.TabIndex = 0x5b;
            this.button10.Text = "---";
            this.button10.UseVisualStyleBackColor = false;
            this.label1.AutoSize = true;
            this.label1.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.label1.ForeColor = Color.Lime;
            this.label1.Location = new Point(0x127, 0xa8);
            this.label1.Name = "label1";
            this.label1.Size = new Size(0x13, 13);
            this.label1.TabIndex = 0x4a;
            this.label1.Text = "---";
            this.label53.AutoSize = true;
            this.label53.Location = new Point(0x124, 0x89);
            this.label53.Name = "label53";
            this.label53.Size = new Size(0x4a, 15);
            this.label53.TabIndex = 90;
            this.label53.Text = "Frame Price";
            this.total2.BackColor = Color.AliceBlue;
            this.total2.Enabled = false;
            this.total2.FlatStyle = FlatStyle.Popup;
            this.total2.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.total2.Location = new Point(0x188, 0x85);
            this.total2.Name = "total2";
            this.total2.Size = new Size(0x3e, 0x17);
            this.total2.TabIndex = 0x59;
            this.total2.Text = "---";
            this.total2.UseVisualStyleBackColor = false;
            this.label52.AutoSize = true;
            this.label52.Location = new Point(0x58, 0x89);
            this.label52.Name = "label52";
            this.label52.Size = new Size(0x41, 15);
            this.label52.TabIndex = 0x58;
            this.label52.Text = "Lens Price";
            this.total1.BackColor = Color.AliceBlue;
            this.total1.Enabled = false;
            this.total1.FlatStyle = FlatStyle.Popup;
            this.total1.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.total1.Location = new Point(0xbc, 0x85);
            this.total1.Name = "total1";
            this.total1.Size = new Size(0x3e, 0x17);
            this.total1.TabIndex = 0x57;
            this.total1.Text = "---";
            this.total1.UseVisualStyleBackColor = false;
            this.label51.AutoSize = true;
            this.label51.Location = new Point(0x58, 0xa4);
            this.label51.Name = "label51";
            this.label51.Size = new Size(0x4f, 15);
            this.label51.TabIndex = 0x56;
            this.label51.Text = "Total Amount";
            this.total3.BackColor = Color.AliceBlue;
            this.total3.Enabled = false;
            this.total3.FlatStyle = FlatStyle.Popup;
            this.total3.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.total3.Location = new Point(0xbc, 160);
            this.total3.Name = "total3";
            this.total3.Size = new Size(0x3e, 0x17);
            this.total3.TabIndex = 0x55;
            this.total3.Text = "---";
            this.total3.UseVisualStyleBackColor = false;
            this.ll6.BackColor = Color.PaleGreen;
            this.ll6.Enabled = false;
            this.ll6.FlatStyle = FlatStyle.Popup;
            this.ll6.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.ll6.Location = new Point(0x1c8, 90);
            this.ll6.Name = "ll6";
            this.ll6.Size = new Size(0x3e, 0x17);
            this.ll6.TabIndex = 0x54;
            this.ll6.Text = "---";
            this.ll6.UseVisualStyleBackColor = false;
            this.lr6.BackColor = Color.PaleGreen;
            this.lr6.Enabled = false;
            this.lr6.FlatStyle = FlatStyle.Popup;
            this.lr6.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lr6.Location = new Point(0x1c8, 0x41);
            this.lr6.Name = "lr6";
            this.lr6.Size = new Size(0x3e, 0x17);
            this.lr6.TabIndex = 0x53;
            this.lr6.Text = "---";
            this.lr6.UseVisualStyleBackColor = false;
            this.ll5.BackColor = Color.AliceBlue;
            this.ll5.Enabled = false;
            this.ll5.FlatStyle = FlatStyle.Popup;
            this.ll5.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.ll5.Location = new Point(0x184, 90);
            this.ll5.Name = "ll5";
            this.ll5.Size = new Size(0x3e, 0x17);
            this.ll5.TabIndex = 0x52;
            this.ll5.Text = "---";
            this.ll5.UseVisualStyleBackColor = false;
            this.lr5.BackColor = Color.AliceBlue;
            this.lr5.Enabled = false;
            this.lr5.FlatStyle = FlatStyle.Popup;
            this.lr5.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lr5.Location = new Point(0x184, 0x41);
            this.lr5.Name = "lr5";
            this.lr5.Size = new Size(0x3e, 0x17);
            this.lr5.TabIndex = 0x51;
            this.lr5.Text = "---";
            this.lr5.UseVisualStyleBackColor = false;
            this.ll4.BackColor = Color.Pink;
            this.ll4.Enabled = false;
            this.ll4.FlatStyle = FlatStyle.Popup;
            this.ll4.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.ll4.Location = new Point(320, 90);
            this.ll4.Name = "ll4";
            this.ll4.Size = new Size(0x3e, 0x17);
            this.ll4.TabIndex = 80;
            this.ll4.Text = "---";
            this.ll4.UseVisualStyleBackColor = false;
            this.lr4.BackColor = Color.Pink;
            this.lr4.Enabled = false;
            this.lr4.FlatStyle = FlatStyle.Popup;
            this.lr4.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lr4.Location = new Point(320, 0x41);
            this.lr4.Name = "lr4";
            this.lr4.Size = new Size(0x3e, 0x17);
            this.lr4.TabIndex = 0x4f;
            this.lr4.Text = "---";
            this.lr4.UseVisualStyleBackColor = false;
            this.ll3.BackColor = Color.Lavender;
            this.ll3.Enabled = false;
            this.ll3.FlatStyle = FlatStyle.Popup;
            this.ll3.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.ll3.Location = new Point(0xfc, 90);
            this.ll3.Name = "ll3";
            this.ll3.Size = new Size(0x3e, 0x17);
            this.ll3.TabIndex = 0x4e;
            this.ll3.Text = "---";
            this.ll3.UseVisualStyleBackColor = false;
            this.lr3.BackColor = Color.Lavender;
            this.lr3.Enabled = false;
            this.lr3.FlatStyle = FlatStyle.Popup;
            this.lr3.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lr3.Location = new Point(0xfc, 0x41);
            this.lr3.Name = "lr3";
            this.lr3.Size = new Size(0x3e, 0x17);
            this.lr3.TabIndex = 0x4d;
            this.lr3.Text = "---";
            this.lr3.UseVisualStyleBackColor = false;
            this.ll2.BackColor = Color.MintCream;
            this.ll2.Enabled = false;
            this.ll2.FlatStyle = FlatStyle.Popup;
            this.ll2.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.ll2.Location = new Point(0xb8, 90);
            this.ll2.Name = "ll2";
            this.ll2.Size = new Size(0x3e, 0x17);
            this.ll2.TabIndex = 0x4c;
            this.ll2.Text = "---";
            this.ll2.UseVisualStyleBackColor = false;
            this.lr2.BackColor = Color.MintCream;
            this.lr2.Enabled = false;
            this.lr2.FlatStyle = FlatStyle.Popup;
            this.lr2.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lr2.Location = new Point(0xb8, 0x41);
            this.lr2.Name = "lr2";
            this.lr2.Size = new Size(0x3e, 0x17);
            this.lr2.TabIndex = 0x4b;
            this.lr2.Text = "---";
            this.lr2.UseVisualStyleBackColor = false;
            this.ll1.BackColor = Color.MistyRose;
            this.ll1.Enabled = false;
            this.ll1.FlatStyle = FlatStyle.Popup;
            this.ll1.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.ll1.Location = new Point(0x71, 90);
            this.ll1.Name = "ll1";
            this.ll1.Size = new Size(0x3e, 0x17);
            this.ll1.TabIndex = 0x4a;
            this.ll1.Text = "---";
            this.ll1.UseVisualStyleBackColor = false;
            this.lr1.BackColor = Color.MistyRose;
            this.lr1.Enabled = false;
            this.lr1.FlatStyle = FlatStyle.Popup;
            this.lr1.Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lr1.Location = new Point(0x71, 0x41);
            this.lr1.Name = "lr1";
            this.lr1.Size = new Size(0x3e, 0x17);
            this.lr1.TabIndex = 0x49;
            this.lr1.Text = "---";
            this.lr1.UseVisualStyleBackColor = false;
            this.fm_type.AutoSize = true;
            this.fm_type.Font = new Font("Garamond", 11.25f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.fm_type.ForeColor = SystemColors.MenuHighlight;
            this.fm_type.Location = new Point(0x264, 100);
            this.fm_type.Name = "fm_type";
            this.fm_type.Size = new Size(0x17, 0x11);
            this.fm_type.TabIndex = 0x47;
            this.fm_type.Text = "---";
            this.lb_side.AutoSize = true;
            this.lb_side.Font = new Font("Garamond", 11.25f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lb_side.ForeColor = SystemColors.MenuHighlight;
            this.lb_side.Location = new Point(0x264, 80);
            this.lb_side.Name = "lb_side";
            this.lb_side.Size = new Size(0x17, 0x11);
            this.lb_side.TabIndex = 70;
            this.lb_side.Text = "---";
            this.lb_for.AutoSize = true;
            this.lb_for.Font = new Font("Garamond", 11.25f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lb_for.ForeColor = SystemColors.MenuHighlight;
            this.lb_for.Location = new Point(0x264, 0x36);
            this.lb_for.Name = "lb_for";
            this.lb_for.Size = new Size(0x17, 0x11);
            this.lb_for.TabIndex = 0x45;
            this.lb_for.Text = "---";
            this.lb_type.AutoSize = true;
            this.lb_type.Font = new Font("Garamond", 11.25f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.lb_type.ForeColor = SystemColors.MenuHighlight;
            this.lb_type.Location = new Point(0x264, 0x1b);
            this.lb_type.Name = "lb_type";
            this.lb_type.Size = new Size(0x17, 0x11);
            this.lb_type.TabIndex = 0x44;
            this.lb_type.Text = "---";
            this.label58.AutoSize = true;
            this.label58.Location = new Point(0x216, 80);
            this.label58.Name = "label58";
            this.label58.Size = new Size(0x3e, 15);
            this.label58.TabIndex = 0x36;
            this.label58.Text = "Lens Side";
            this.label59.AutoSize = true;
            this.label59.Location = new Point(0x217, 0x36);
            this.label59.Name = "label59";
            this.label59.Size = new Size(0x37, 15);
            this.label59.TabIndex = 0x34;
            this.label59.Text = "Lens For";
            this.label60.AutoSize = true;
            this.label60.Location = new Point(0x217, 0x1b);
            this.label60.Name = "label60";
            this.label60.Size = new Size(0x3f, 15);
            this.label60.TabIndex = 0x33;
            this.label60.Text = "Lens Type";
            this.label62.AutoSize = true;
            this.label62.Location = new Point(0x217, 100);
            this.label62.Name = "label62";
            this.label62.Size = new Size(0x48, 15);
            this.label62.TabIndex = 0x35;
            this.label62.Text = "Frame Type";
            this.label49.AutoSize = true;
            this.label49.Location = new Point(8, 0x5d);
            this.label49.Name = "label49";
            this.label49.Size = new Size(0x34, 15);
            this.label49.TabIndex = 50;
            this.label49.Text = "OS(Left)";
            this.label48.AutoSize = true;
            this.label48.Location = new Point(6, 0x2a);
            this.label48.Name = "label48";
            this.label48.Size = new Size(0x37, 15);
            this.label48.TabIndex = 0x3d;
            this.label48.Text = "Distance";
            this.label50.AutoSize = true;
            this.label50.Location = new Point(8, 0x45);
            this.label50.Name = "label50";
            this.label50.Size = new Size(0x3e, 15);
            this.label50.TabIndex = 0x31;
            this.label50.Text = "OD(Right)";
            this.label34.AutoSize = true;
            this.label34.Location = new Point(0x1c5, 0x24);
            this.label34.Name = "label34";
            this.label34.Size = new Size(0x17, 15);
            this.label34.TabIndex = 0x36;
            this.label34.Text = "VN";
            this.label39.AutoSize = true;
            this.label39.Location = new Point(110, 0x24);
            this.label39.Name = "label39";
            this.label39.Size = new Size(0x23, 15);
            this.label39.TabIndex = 0x31;
            this.label39.Text = "SPH.";
            this.label35.AutoSize = true;
            this.label35.Location = new Point(0x181, 0x1b);
            this.label35.Name = "label35";
            this.label35.Size = new Size(0x2b, 30);
            this.label35.TabIndex = 0x35;
            this.label35.Text = "NEAR \r\nADD";
            this.label38.AutoSize = true;
            this.label38.Location = new Point(0xb5, 0x24);
            this.label38.Name = "label38";
            this.label38.Size = new Size(0x20, 15);
            this.label38.TabIndex = 50;
            this.label38.Text = "CYL.";
            this.label36.AutoSize = true;
            this.label36.Location = new Point(0x13d, 0x24);
            this.label36.Name = "label36";
            this.label36.Size = new Size(0x17, 15);
            this.label36.TabIndex = 0x34;
            this.label36.Text = "VD";
            this.label37.AutoSize = true;
            this.label37.Location = new Point(0xf9, 0x24);
            this.label37.Name = "label37";
            this.label37.Size = new Size(0x21, 15);
            this.label37.TabIndex = 0x33;
            this.label37.Text = "AXIS";
            this.groupBox5.Location = new Point(9, 0x31);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new Size(0x16c, 0x86);
            this.groupBox5.TabIndex = 0x24;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Left Eye";
            this.lb_contact.AutoSize = true;
            this.lb_contact.Location = new Point(0x1d1, 0x11);
            this.lb_contact.Name = "lb_contact";
            this.lb_contact.Size = new Size(0x13, 15);
            this.lb_contact.TabIndex = 0x23;
            this.lb_contact.Text = "---";
            this.lb_name.AutoSize = true;
            this.lb_name.Location = new Point(0xfe, 0x11);
            this.lb_name.Name = "lb_name";
            this.lb_name.Size = new Size(0x13, 15);
            this.lb_name.TabIndex = 0x22;
            this.lb_name.Text = "---";
            this.label30.AutoSize = true;
            this.label30.Location = new Point(0x188, 0x11);
            this.label30.Name = "label30";
            this.label30.Size = new Size(0x43, 15);
            this.label30.TabIndex = 0x11;
            this.label30.Text = "Contact No";
            this.label31.AutoSize = true;
            this.label31.Location = new Point(0xcf, 0x11);
            this.label31.Name = "label31";
            this.label31.Size = new Size(0x29, 15);
            this.label31.TabIndex = 0x10;
            this.label31.Text = "Name";
            this.saveFileDialog1.DefaultExt = "bak";
            this.saveFileDialog1.FileName = "OpticalDb";
            this.saveFileDialog1.InitialDirectory = @"C:\OpticalShop\Backups";
            this.saveFileDialog1.FileOk += new CancelEventHandler(this.saveFileDialog1_FileOk);
            this.openFileDialog1.DefaultExt = "bak";
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.InitialDirectory = @"C:\OpticalShop\Backups";
            this.openFileDialog1.FileOk += new CancelEventHandler(this.openFileDialog1_FileOk);
            base.AcceptButton = this.button1;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x4f7, 0x2cf);
            base.Controls.Add(this.groupBox4);
            base.Controls.Add(this.groupBox3);
            base.Controls.Add(this.tabControl1);
            base.Controls.Add(this.menuStrip1);
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.KeyPreview = true;
            base.MainMenuStrip = this.menuStrip1;
            base.Name = "Form1";
            this.Text = "Eye+";
            base.FormClosing += new FormClosingEventHandler(this.Form1_FormClosing);
            base.FormClosed += new FormClosedEventHandler(this.Form1_FormClosed);
            base.Load += new EventHandler(this.Form1_Load);
            base.PreviewKeyDown += new PreviewKeyDownEventHandler(this.Form1_PreviewKeyDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((ISupportInitialize) this.dataGridView2).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((ISupportInitialize) this.dataGridView1).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        private void lens4_TextChanged(object sender, EventArgs e)
        {
            float num = (Convert.ToInt32("0" + this.f2.Text) + Convert.ToInt32("0" + this.lens4.Text)) + Convert.ToInt32("0" + this.p2.Text);
            this.p3.Text = this.p4.Text = num.ToString();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            File.Copy(this.openFileDialog1.FileName, "C:/OpticalShop/Database/OpticalDb.mdb", true);
            MessageBox.Show("Database Has been Imported!!");
            this.Refresh();
        }

        private void p2_TextChanged(object sender, EventArgs e)
        {
            float num = (Convert.ToInt32("0" + this.f2.Text) + Convert.ToInt32("0" + this.lens4.Text)) + Convert.ToInt32("0" + this.p2.Text);
            this.p3.Text = this.p4.Text = num.ToString();
        }

        private void p3_TextChanged(object sender, EventArgs e)
        {
        }

        private void p3_TextChanged_1(object sender, EventArgs e)
        {
            this.p5.Text = ((float) (Convert.ToInt32("0" + this.p4.Text) - Convert.ToInt32("0" + this.p3.Text))).ToString();
            this.pen_status.Text = (Convert.ToInt32(this.p5.Text) > 0) ? "Pending" : "Paid";
            this.label1.Text = this.pen_status.Text;
            this.button10.Text = this.p3.Text;
        }

        private void p4_TextChanged(object sender, EventArgs e)
        {
        }

        private void p4_TextChanged_1(object sender, EventArgs e)
        {
            this.p5.Text = ((float) (Convert.ToInt32("0" + this.p4.Text) - Convert.ToInt32("0" + this.p3.Text))).ToString();
            if (Convert.ToInt32(this.p5.Text) <= 0)
            {
                this.pen_status.Text = "Paid";
            }
            else
            {
                this.pen_status.Text = "Pending";
            }
        }

        private void p5_TextChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(this.p5.Text) <= 0)
            {
                this.pen_status.Text = "Paid";
            }
            else
            {
                this.pen_status.Text = "Pending";
            }
        }

        private void pen_status_TextChanged(object sender, EventArgs e)
        {
            this.label1.Text = this.pen_status.Text;
            this.button10.Text = this.p3.Text;
        }

        public void Refresh()
        {
            try
            {
                this.conn = new OleDbConnection();
                this.conn.ConnectionString = this.connstr;
                this.conn.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter {
                    SelectCommand = new OleDbCommand("SELECT  `Ex_by`,`name`, `Contact`,`Order_date`,`Age`, `Addr`, `Gender`,`ID`,`birthdate` FROM `Customer` ORDER BY ID DESC", this.conn)
                };
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                BindingSource source = new BindingSource {
                    DataSource = dataTable
                };
                this.dataGridView1.DataSource = source;
                adapter.Update(dataTable);
            }
            catch (Exception exception1)
            {
                string message = exception1.Message;
                MessageBox.Show(message ?? "");
            }
            this.dataGridView1.Columns[1].Width = 200;
            this.dataGridView1.Columns[4].Width = 40;
            this.dataGridView1.Columns[7].Width = 100;
        }

        public void Refresh_Pre()
        {
            try
            {
                this.conn = new OleDbConnection();
                this.conn.ConnectionString = this.connstr;
                this.conn.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter {
                    SelectCommand = new OleDbCommand("SELECT Pre_ID,Order_date,ID FROM `Lens_Pre` where ID=" + this.cust_id.Text + "  ORDER BY Pre_ID DESC", this.conn)
                };
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                BindingSource source = new BindingSource {
                    DataSource = dataTable
                };
                this.dataGridView2.DataSource = source;
                adapter.Update(dataTable);
            }
            catch (Exception exception1)
            {
                string message = exception1.Message;
                MessageBox.Show(message ?? "");
            }
        }

        private void right1_KeyUp(object sender, KeyEventArgs e)
        {
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            File.Copy("C:/OpticalShop/Database/OpticalDb.mdb", this.saveFileDialog1.FileName, true);
            MessageBox.Show("Database Has been backup at " + this.saveFileDialog1.FileName + " !");
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void Search_txt_TextChanged(object sender, EventArgs e)
        {
            try
            {
                this.conn = new OleDbConnection();
                OleDbCommand command = new OleDbCommand("SELECT `Ex_by`,`name`, `Contact`,`Order_date`,`Age`, `Addr` , `Gender`,`ID`,`Ref_no`,`birthdate` FROM `Customer` WHERE name LIKE '%" + this.Search_txt.Text + "%'  ORDER BY ID DESC", this.conn);
                this.conn.ConnectionString = this.connstr;
                this.conn.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                if (this.Search_by.Text == "Pending Bill")
                {
                    command = new OleDbCommand("SELECT `Ex_by`,`name`, `Contact`,`Order_date`,`Age`, `Addr` , `Gender`,`ID`,`Ref_no`,`birthdate` FROM `Customer` WHERE ID IN (SELECT ID FROM Lens_Pre WHERE Pending_status LIKE '%" + this.Search_txt.Text + "%')", this.conn);
                }
                else if (this.Search_by.Text == "Examined By")
                {
                    command = new OleDbCommand("SELECT `Ex_by`,`name`, `Contact`,`Order_date`,`Age`, `Addr` , `Gender`,`ID`,`Ref_no`,`birthdate` FROM `Customer` WHERE Ex_by LIKE '%" + this.Search_txt.Text + "%'  ORDER BY ID DESC", this.conn);
                }
                else if (this.Search_by.Text == "Contact no")
                {
                    command = new OleDbCommand("SELECT `Ex_by`,`name`, `Contact`,`Order_date`,`Age`, `Addr` , `Gender`,`ID`,`Ref_no`,`birthdate` FROM `Customer` WHERE Contact LIKE '%" + this.Search_txt.Text + "%'  ORDER BY ID DESC", this.conn);
                }
                else if (this.Search_by.Text == "Order date")
                {
                    command = new OleDbCommand("SELECT `Ex_by`,`name`, `Contact`,`Order_date`,`Age`, `Addr` , `Gender`,`ID`,`Ref_no`,`birthdate` FROM `Customer` WHERE Order_date LIKE '%" + this.Search_txt.Text + "%'  ORDER BY ID DESC", this.conn);
                }
                adapter.SelectCommand = command;
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                BindingSource source = new BindingSource {
                    DataSource = dataTable
                };
                this.dataGridView1.DataSource = source;
                adapter.Update(dataTable);
            }
            catch (Exception exception1)
            {
                string message = exception1.Message;
                MessageBox.Show(message ?? "");
            }
        }

        private void t6_TextChanged(object sender, EventArgs e)
        {
            if (this.t6.Text.Length > 10)
            {
                MessageBox.Show("Contact Number Should be less than 10 digit");
            }
        }
    }
}

