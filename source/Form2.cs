using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;


namespace PPManager
{
	/// <summary>
	/// Summary description for Form2.
	/// </summary>
	public class frmAbout : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblText1;
		private System.Windows.Forms.PictureBox picAbout;
		private System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Label lblText2;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblAuthor;
		private System.Windows.Forms.PictureBox picAuthor;
		private System.ComponentModel.IContainer components;

		Screen [] m_Screen = Screen.AllScreens;
		Rectangle m_rect = new Rectangle();
		int m_iLeft, m_iCenter;

		public frmAbout()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAbout));
            this.lblText1 = new System.Windows.Forms.Label();
            this.picAbout = new System.Windows.Forms.PictureBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblText2 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblAuthor = new System.Windows.Forms.Label();
            this.picAuthor = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.picAbout)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picAuthor)).BeginInit();
            this.SuspendLayout();
            // 
            // lblText1
            // 
            this.lblText1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblText1.Location = new System.Drawing.Point(32, 64);
            this.lblText1.Name = "lblText1";
            this.lblText1.Size = new System.Drawing.Size(248, 40);
            this.lblText1.TabIndex = 0;
            this.lblText1.Text = "Copyright (C) Andrew Leong 2009. All rights reserved.";
            this.lblText1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblText1.Click += new System.EventHandler(this.lblText1_Click);
            // 
            // picAbout
            // 
            this.picAbout.Image = ((System.Drawing.Image)(resources.GetObject("picAbout.Image")));
            this.picAbout.Location = new System.Drawing.Point(16, 16);
            this.picAbout.Name = "picAbout";
            this.picAbout.Size = new System.Drawing.Size(32, 32);
            this.picAbout.TabIndex = 1;
            this.picAbout.TabStop = false;
            // 
            // lblTitle
            // 
            this.lblTitle.Font = new System.Drawing.Font("Palatino Linotype", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Blue;
            this.lblTitle.Location = new System.Drawing.Point(56, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(208, 23);
            this.lblTitle.TabIndex = 2;
            this.lblTitle.Text = "Power Point Manager v1.0";
            // 
            // lblText2
            // 
            this.lblText2.Font = new System.Drawing.Font("Dotum", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblText2.Location = new System.Drawing.Point(32, 120);
            this.lblText2.Name = "lblText2";
            this.lblText2.Size = new System.Drawing.Size(224, 144);
            this.lblText2.TabIndex = 3;
            this.lblText2.Text = resources.GetString("lblText2.Text");
            // 
            // btnOK
            // 
            this.btnOK.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.ForeColor = System.Drawing.Color.Black;
            this.btnOK.Location = new System.Drawing.Point(445, 291);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "&Close";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblAuthor
            // 
            this.lblAuthor.Font = new System.Drawing.Font("Dotum", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAuthor.Location = new System.Drawing.Point(312, 128);
            this.lblAuthor.Name = "lblAuthor";
            this.lblAuthor.Size = new System.Drawing.Size(208, 136);
            this.lblAuthor.TabIndex = 6;
            this.lblAuthor.Text = resources.GetString("lblAuthor.Text");
            // 
            // picAuthor
            // 
            this.picAuthor.Image = ((System.Drawing.Image)(resources.GetObject("picAuthor.Image")));
            this.picAuthor.Location = new System.Drawing.Point(320, 16);
            this.picAuthor.Name = "picAuthor";
            this.picAuthor.Size = new System.Drawing.Size(136, 96);
            this.picAuthor.TabIndex = 8;
            this.picAuthor.TabStop = false;
            // 
            // frmAbout
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(544, 326);
            this.Controls.Add(this.picAuthor);
            this.Controls.Add(this.lblAuthor);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblText2);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.picAbout);
            this.Controls.Add(this.lblText1);
            this.Name = "frmAbout";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PPManager v1.0";
            this.Load += new System.EventHandler(this.frmAbout_Load);
            ((System.ComponentModel.ISupportInitialize)(this.picAbout)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picAuthor)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
			Close();
		}

		private void frmAbout_Load(object sender, System.EventArgs e)
		{
            //m_rect = m_Screen[0].Bounds;

            //m_iCenter = m_rect.Width / 2;
            //m_iLeft = m_iCenter - 150;

            //this.Left = m_iLeft;

            //this.Size = new System.Drawing.Size(300, 360);
		}

		private void btnAuthor_Click(object sender, System.EventArgs e)
		{
            //m_iLeft = m_iCenter - 276;
			
            //this.Left = m_iLeft;

            //this.Size = new System.Drawing.Size(552, 360);
		}

		private void btnBack_Click(object sender, System.EventArgs e)
		{
            //m_iLeft = m_iCenter - 150;

            //this.Left = m_iLeft;

            //this.Size = new System.Drawing.Size(300, 360);
		}

        private void lblText1_Click(object sender, EventArgs e)
        {

        }
	}
}
