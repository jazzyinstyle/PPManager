using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace PPManager
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class frmMain : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ListView lstList;
		private System.Windows.Forms.OpenFileDialog opnFile;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtDir;
		private System.Windows.Forms.Button btnDir;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnRemove;
		private System.Windows.Forms.Button btnUp;
		private System.Windows.Forms.Button btnDown;
		private System.Windows.Forms.TextBox txtSaveAs;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.NotifyIcon ntfIcon;
		public System.Windows.Forms.ContextMenu mnuMnu;
		private System.Windows.Forms.MenuItem mnuShow;
		private System.Windows.Forms.Button btnRun;
		private System.Windows.Forms.Button btnExit;
		private System.Windows.Forms.MenuItem mnuAbout;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.ColumnHeader colHeader;
		private System.Windows.Forms.MenuItem mnuExit;

		ToolTip m_tooltip;

		public frmMain()
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
				if (components != null) 
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmMain));
			this.lstList = new System.Windows.Forms.ListView();
			this.colHeader = new System.Windows.Forms.ColumnHeader();
			this.opnFile = new System.Windows.Forms.OpenFileDialog();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDir = new System.Windows.Forms.TextBox();
			this.btnDir = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSaveAs = new System.Windows.Forms.TextBox();
			this.btnRemove = new System.Windows.Forms.Button();
			this.btnUp = new System.Windows.Forms.Button();
			this.btnDown = new System.Windows.Forms.Button();
			this.ntfIcon = new System.Windows.Forms.NotifyIcon(this.components);
			this.mnuMnu = new System.Windows.Forms.ContextMenu();
			this.mnuShow = new System.Windows.Forms.MenuItem();
			this.mnuAbout = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.mnuExit = new System.Windows.Forms.MenuItem();
			this.btnRun = new System.Windows.Forms.Button();
			this.btnExit = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// lstList
			// 
			this.lstList.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.lstList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																					  this.colHeader});
			this.lstList.FullRowSelect = true;
			this.lstList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstList.HideSelection = false;
			this.lstList.Location = new System.Drawing.Point(16, 112);
			this.lstList.MultiSelect = false;
			this.lstList.Name = "lstList";
			this.lstList.Size = new System.Drawing.Size(432, 160);
			this.lstList.TabIndex = 1;
			this.lstList.View = System.Windows.Forms.View.List;
			// 
			// colHeader
			// 
			this.colHeader.Text = "Presentation Order";
			this.colHeader.Width = 500;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Presentation:";
			// 
			// txtDir
			// 
			this.txtDir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtDir.Location = new System.Drawing.Point(88, 24);
			this.txtDir.Name = "txtDir";
			this.txtDir.Size = new System.Drawing.Size(328, 20);
			this.txtDir.TabIndex = 3;
			this.txtDir.Text = "";
			this.txtDir.MouseEnter += new System.EventHandler(this.txtDir_MouseEnter);
			// 
			// btnDir
			// 
			this.btnDir.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnDir.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnDir.Image = ((System.Drawing.Image)(resources.GetObject("btnDir.Image")));
			this.btnDir.Location = new System.Drawing.Point(432, 24);
			this.btnDir.Name = "btnDir";
			this.btnDir.Size = new System.Drawing.Size(40, 32);
			this.btnDir.TabIndex = 4;
			this.btnDir.Click += new System.EventHandler(this.btnDir_Click);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(16, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 23);
			this.label2.TabIndex = 7;
			this.label2.Text = "Save As:";
			// 
			// txtSaveAs
			// 
			this.txtSaveAs.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtSaveAs.Location = new System.Drawing.Point(88, 56);
			this.txtSaveAs.Name = "txtSaveAs";
			this.txtSaveAs.Size = new System.Drawing.Size(328, 20);
			this.txtSaveAs.TabIndex = 8;
			this.txtSaveAs.Text = "";
			this.txtSaveAs.MouseEnter += new System.EventHandler(this.txtSaveAs_MouseEnter);
			// 
			// btnRemove
			// 
			this.btnRemove.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnRemove.Image = ((System.Drawing.Image)(resources.GetObject("btnRemove.Image")));
			this.btnRemove.Location = new System.Drawing.Point(456, 192);
			this.btnRemove.Name = "btnRemove";
			this.btnRemove.Size = new System.Drawing.Size(32, 23);
			this.btnRemove.TabIndex = 9;
			this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
			// 
			// btnUp
			// 
			this.btnUp.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnUp.Image = ((System.Drawing.Image)(resources.GetObject("btnUp.Image")));
			this.btnUp.Location = new System.Drawing.Point(456, 112);
			this.btnUp.Name = "btnUp";
			this.btnUp.Size = new System.Drawing.Size(32, 23);
			this.btnUp.TabIndex = 10;
			this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
			// 
			// btnDown
			// 
			this.btnDown.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnDown.Image = ((System.Drawing.Image)(resources.GetObject("btnDown.Image")));
			this.btnDown.Location = new System.Drawing.Point(456, 152);
			this.btnDown.Name = "btnDown";
			this.btnDown.Size = new System.Drawing.Size(32, 23);
			this.btnDown.TabIndex = 11;
			this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
			// 
			// ntfIcon
			// 
			this.ntfIcon.ContextMenu = this.mnuMnu;
			this.ntfIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("ntfIcon.Icon")));
			this.ntfIcon.Text = "PPManager";
			this.ntfIcon.Visible = true;
			this.ntfIcon.DoubleClick += new System.EventHandler(this.ntfIcon_DoubleClick);
			// 
			// mnuMnu
			// 
			this.mnuMnu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																				   this.mnuShow,
																				   this.mnuAbout,
																				   this.menuItem1,
																				   this.mnuExit});
			// 
			// mnuShow
			// 
			this.mnuShow.Index = 0;
			this.mnuShow.Text = "&Show";
			this.mnuShow.Click += new System.EventHandler(this.mnuShow_Click);
			// 
			// mnuAbout
			// 
			this.mnuAbout.Index = 1;
			this.mnuAbout.Text = "&About";
			this.mnuAbout.Click += new System.EventHandler(this.mnuAbout_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "-";
			// 
			// mnuExit
			// 
			this.mnuExit.Index = 3;
			this.mnuExit.Text = "&Exit";
			this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
			// 
			// btnRun
			// 
			this.btnRun.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnRun.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnRun.Image = ((System.Drawing.Image)(resources.GetObject("btnRun.Image")));
			this.btnRun.Location = new System.Drawing.Point(216, 280);
			this.btnRun.Name = "btnRun";
			this.btnRun.Size = new System.Drawing.Size(72, 48);
			this.btnRun.TabIndex = 12;
			this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
			// 
			// btnExit
			// 
			this.btnExit.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
			this.btnExit.Location = new System.Drawing.Point(448, 288);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(40, 40);
			this.btnExit.TabIndex = 13;
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// frmMain
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(496, 342);
			this.Controls.Add(this.btnExit);
			this.Controls.Add(this.btnRun);
			this.Controls.Add(this.btnDown);
			this.Controls.Add(this.btnUp);
			this.Controls.Add(this.btnRemove);
			this.Controls.Add(this.txtSaveAs);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnDir);
			this.Controls.Add(this.txtDir);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lstList);
			this.MaximizeBox = false;
			this.Name = "frmMain";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Power Point Manager v1.0";
			this.Resize += new System.EventHandler(this.frmMain_Resize);
			this.Load += new System.EventHandler(this.frmMain_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frmMain());
		}

		private void btnRun_Click(object sender, System.EventArgs e)
		{
			if(lstList.Items.Count <= 0)
				return;

			this.Hide();
			ShowPresentation();
			GC.Collect();
		}

		private void ShowPresentation()
		{
			String strDest = Environment.CurrentDirectory + "\\" + txtSaveAs.Text + ".ppt";

			PowerPoint.Application objApp;
			PowerPoint.Presentations objPresSet;
			PowerPoint._Presentation objPres;
			PowerPoint.SlideShowSettings objSSS;	
			PowerPoint.SlideShowWindows objSSWs;
			
			//Create a new presentation based on a template.
			objApp = new PowerPoint.Application();
			objApp.Visible = MsoTriState.msoTrue;
			objPresSet = objApp.Presentations;

			int iCurCnt = 1;
			int iSlideCnt = 0;

			objPres = objPresSet.Add(Microsoft.Office.Core.MsoTriState.msoFalse);

			for(int i = 0; i < lstList.Items.Count; i++)
			{
				PowerPoint._Presentation objPresOpen;

				objPresOpen = objPresSet.Open(lstList.Items[i].Text, MsoTriState.msoTrue, MsoTriState.msoTrue, 
					MsoTriState.msoFalse);
				iSlideCnt = objPresOpen.Slides.Count;
			
				objPres.Slides.InsertFromFile(lstList.Items[i].Text, iCurCnt - 1, 1, iSlideCnt);
				
				for(int j = 1, iCount = iCurCnt; j <= iSlideCnt; j++, iCount++)
				{
					PowerPoint.Slide slide = objPres.Slides.Item(iCount);
					slide.Design = objPresOpen.Slides.Item(j).Design;
					slide.ColorScheme = objPresOpen.Slides.Item(j).ColorScheme;
					slide.FollowMasterBackground = MsoTriState.msoTrue;
					slide.DisplayMasterShapes = MsoTriState.msoTrue;
				}

				objPresOpen.Close();

				iCurCnt += iSlideCnt;
			}
			
			objPres.SaveAs(strDest, PowerPoint.PpSaveAsFileType.ppSaveAsPresentation,
				MsoTriState.msoTrue);

			objPres.Close();


			objPres = objPresSet.Open(strDest, MsoTriState.msoTrue, MsoTriState.msoTrue, 
				MsoTriState.msoTrue);
			
			//Prevent Office Assistant from displaying alert messages:
			objApp.Assistant.On = false;
			objApp.Assistant.Visible = false;
		

			//Run the Slide show from slides 1 thru 3.
			objSSS = objPres.SlideShowSettings;
			objSSS.StartingSlide = 1;
			objSSS.EndingSlide = iCurCnt - 1;
			objSSS.Run();

			//Wait for the slide show to end.
			objSSWs = objApp.SlideShowWindows;
			while(objSSWs.Count >= 1) 
				System.Threading.Thread.Sleep(100);
		
			//Close the presentation without saving changes and quit PowerPoint.
			objPres.Close();
			objApp.Quit();
		}

		private void frmMain_Load(object sender, System.EventArgs e)
		{
			ntfIcon.Visible = true;
			lstList.Items.Clear();
			txtDir.Text = "";
			txtSaveAs.Text = "Temp";

			m_tooltip = new ToolTip();
			m_tooltip.InitialDelay = 500;
			m_tooltip.SetToolTip(btnDir, "Browse to Add");
			m_tooltip.SetToolTip(btnUp, "Up");
			m_tooltip.SetToolTip(btnDown, "Down");
			m_tooltip.SetToolTip(btnRemove, "Remove");
			m_tooltip.SetToolTip(btnRun, "Run");
			m_tooltip.SetToolTip(btnExit, "Exit");
		}

		private void btnDir_Click(object sender, System.EventArgs e)
		{
			opnFile.Title = "Presentation files";
			opnFile.CheckFileExists = true;
			opnFile.CheckPathExists = true;
			opnFile.Filter = "(Ppt files (*.ppt)|*.ppt";
			
			if(opnFile.ShowDialog(this) == DialogResult.OK)
			{
				txtDir.Text = opnFile.FileName;

				if(txtDir.Text.Trim().Length > 0)
					lstList.Items.Add(txtDir.Text);
			}
			
		}

		private void btnExit_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btnUp_Click(object sender, System.EventArgs e)
		{
			if(lstList.SelectedIndices.Count <= 0)
				return;

			String strTemp;
			int iSel = lstList.FocusedItem.Index;

			if(iSel > 0)
			{
				strTemp = lstList.Items[iSel - 1].Text;
				lstList.Items[iSel - 1].Text = lstList.Items[iSel].Text;
				lstList.Items[iSel].Text = strTemp;

				lstList.Items[iSel - 1].Focused = true;
				lstList.Items[iSel - 1].Selected = true;
			}
		}

		private void btnDown_Click(object sender, System.EventArgs e)
		{
			if(lstList.SelectedIndices.Count <= 0)
				return;

			String strTemp;
			int iSel = lstList.FocusedItem.Index;

			if(iSel < lstList.Items.Count - 1)
			{
				strTemp = lstList.Items[iSel + 1].Text;
				lstList.Items[iSel + 1].Text = lstList.Items[iSel].Text;
				lstList.Items[iSel].Text = strTemp;

				lstList.Items[iSel + 1].Focused = true;
				lstList.Items[iSel + 1].Selected = true;
			}
		}

		private void btnRemove_Click(object sender, System.EventArgs e)
		{
			if(lstList.SelectedIndices.Count <= 0)
				return;

			if(lstList.Items.Count > 0)
			{
				int iSel = lstList.FocusedItem.Index;
				
				if(iSel >= 0)
					lstList.Items[iSel].Remove();
			}
		}

		private void mnuShow_Click(object sender, System.EventArgs e)
		{
			ShowForm();
		}

		private void mnuExit_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void mnuAbout_Click(object sender, System.EventArgs e)
		{
			this.Hide();
			frmAbout About = new frmAbout();
			
			if(About.ShowDialog() == DialogResult.OK)
			{
				ShowForm();
			}
		}

		private void ntfIcon_DoubleClick(object sender, System.EventArgs e)
		{
			ShowForm();
		}

		private void frmMain_Resize(object sender, System.EventArgs e)
		{
			if(this.WindowState == System.Windows.Forms.FormWindowState.Minimized)
			{
				this.Hide();
			}
		}

		private void ShowForm()
		{
			this.Show();
			this.BringToFront();
			this.WindowState = System.Windows.Forms.FormWindowState.Normal;
		}

		private void txtDir_MouseEnter(object sender, System.EventArgs e)
		{
			m_tooltip.SetToolTip(txtDir, txtDir.Text);
		}

		private void txtSaveAs_MouseEnter(object sender, System.EventArgs e)
		{
			m_tooltip.SetToolTip(txtSaveAs, txtSaveAs.Text);
		}
	}
}
