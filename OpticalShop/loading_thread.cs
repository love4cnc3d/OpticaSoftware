namespace OpticalShop
{
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;

    public class loading_thread : Form
    {
        private Timer tmr;
        private IContainer components;
        private Timer timer1;
        private PictureBox pictureBox1;
        private PictureBox pictureBox2;
        private Timer timer2;
        private Label label1;

        public loading_thread()
        {
            this.InitializeComponent();
            this.pictureBox2.Hide();
            this.label1.Hide();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new Container();
            ComponentResourceManager manager = new ComponentResourceManager(typeof(loading_thread));
            this.timer1 = new Timer(this.components);
            this.pictureBox1 = new PictureBox();
            this.pictureBox2 = new PictureBox();
            this.timer2 = new Timer(this.components);
            this.label1 = new Label();
            ((ISupportInitialize) this.pictureBox1).BeginInit();
            ((ISupportInitialize) this.pictureBox2).BeginInit();
            base.SuspendLayout();
            this.timer1.Tick += new EventHandler(this.timer1_Tick);
            this.pictureBox1.Image = (Image) manager.GetObject("pictureBox1.Image");
            this.pictureBox1.Location = new Point(-1, -1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new Size(0x284, 350);
            this.pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox2.Image = (Image) manager.GetObject("pictureBox2.Image");
            this.pictureBox2.Location = new Point(-1, -1);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new Size(0x284, 350);
            this.pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 1;
            this.pictureBox2.TabStop = false;
            this.timer2.Tick += new EventHandler(this.timer2_Tick);
            this.label1.AutoSize = true;
            this.label1.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.label1.ForeColor = SystemColors.ControlDarkDark;
            this.label1.Location = new Point(12, 0x12b);
            this.label1.Name = "label1";
            this.label1.Size = new Size(0xc9, 0x27);
            this.label1.TabIndex = 2;
            this.label1.Text = "17-08-2017 @Designed by Kiran Badave\r\nContact: +91-9673036719\r\nproject.by.kiran@gmail.com\r\n";
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x281, 0x157);
            base.Controls.Add(this.label1);
            base.Controls.Add(this.pictureBox2);
            base.Controls.Add(this.pictureBox1);
            base.FormBorderStyle = FormBorderStyle.None;
            base.Name = "loading_thread";
            base.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "loading_thread";
            base.Load += new EventHandler(this.loading_thread_Load);
            base.Shown += new EventHandler(this.loading_thread_Shown);
            ((ISupportInitialize) this.pictureBox1).EndInit();
            ((ISupportInitialize) this.pictureBox2).EndInit();
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        private void loading_thread_Load(object sender, EventArgs e)
        {
        }

        private void loading_thread_Shown(object sender, EventArgs e)
        {
            this.tmr = new Timer();
            this.tmr.Interval = 0xfa0;
            this.tmr.Start();
            this.tmr.Tick += new EventHandler(this.timer2_Tick);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.tmr.Stop();
            new Form1().Show();
            base.Hide();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            this.tmr.Stop();
            this.label1.Show();
            this.pictureBox2.Show();
            this.pictureBox1.Hide();
            this.tmr = new Timer();
            this.tmr.Interval = 0xfa0;
            this.tmr.Start();
            this.tmr.Tick += new EventHandler(this.timer1_Tick);
        }
    }
}

