using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

[DesignerGenerated]
public sealed class frmAgendaOptions : Form
{
	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTopic")]
	private CheckBox m_A;

	[AccessedThroughProperty("chkAgenda")]
	[CompilerGenerated]
	private CheckBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("GroupBox1")]
	private GroupBox m_A;

	[AccessedThroughProperty("grpBehavior")]
	[CompilerGenerated]
	private GroupBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("CheckBox5")]
	private CheckBox C;

	[AccessedThroughProperty("CheckBox4")]
	[CompilerGenerated]
	private CheckBox D;

	[CompilerGenerated]
	[AccessedThroughProperty("CheckBox3")]
	private CheckBox E;

	[CompilerGenerated]
	[AccessedThroughProperty("Label6")]
	private Label m_A;

	[AccessedThroughProperty("CheckBox6")]
	[CompilerGenerated]
	private CheckBox F;

	[AccessedThroughProperty("RichTextBox1")]
	[CompilerGenerated]
	private RichTextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("RichTextBox2")]
	private RichTextBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("RichTextBox5")]
	private RichTextBox C;

	[AccessedThroughProperty("RichTextBox4")]
	[CompilerGenerated]
	private RichTextBox D;

	[AccessedThroughProperty("RichTextBox3")]
	[CompilerGenerated]
	private RichTextBox E;

	internal virtual Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual Button btnOk
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual CheckBox chkTopic
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EventHandler value2 = A;
			CheckBox checkBox = this.m_A;
			if (checkBox != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				checkBox.CheckedChanged -= value2;
			}
			this.m_A = value;
			checkBox = this.m_A;
			if (checkBox != null)
			{
				checkBox.CheckedChanged += value2;
			}
		}
	}

	internal virtual CheckBox chkAgenda
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EventHandler value2 = A;
			CheckBox checkBox = this.B;
			if (checkBox != null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				checkBox.CheckedChanged -= value2;
			}
			this.B = value;
			checkBox = this.B;
			if (checkBox == null)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				checkBox.CheckedChanged += value2;
				return;
			}
		}
	}

	internal virtual GroupBox GroupBox1
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual GroupBox grpBehavior
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual CheckBox CheckBox5
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	internal virtual CheckBox CheckBox4
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.D = value;
		}
	}

	internal virtual CheckBox CheckBox3
	{
		[CompilerGenerated]
		get
		{
			return this.E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.E = value;
		}
	}

	internal virtual Label Label6
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual CheckBox CheckBox6
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	internal virtual RichTextBox RichTextBox1
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RichTextBox RichTextBox2
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	internal virtual RichTextBox RichTextBox5
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	internal virtual RichTextBox RichTextBox4
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	internal virtual RichTextBox RichTextBox3
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	public frmAgendaOptions(string strPath)
	{
		Application.EnableVisualStyles();
		A();
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (!disposing || this.m_A == null)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_A.Dispose();
				return;
			}
		}
		finally
		{
			base.Dispose(disposing);
		}
	}

	[DebuggerStepThrough]
	private void A()
	{
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmAgendaOptions));
		btnCancel = new Button();
		btnOk = new Button();
		chkTopic = new CheckBox();
		chkAgenda = new CheckBox();
		GroupBox1 = new GroupBox();
		RichTextBox2 = new RichTextBox();
		RichTextBox1 = new RichTextBox();
		grpBehavior = new GroupBox();
		RichTextBox5 = new RichTextBox();
		RichTextBox4 = new RichTextBox();
		RichTextBox3 = new RichTextBox();
		CheckBox5 = new CheckBox();
		CheckBox4 = new CheckBox();
		CheckBox3 = new CheckBox();
		Label6 = new Label();
		CheckBox6 = new CheckBox();
		GroupBox1.SuspendLayout();
		grpBehavior.SuspendLayout();
		SuspendLayout();
		btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnCancel.DialogResult = DialogResult.Cancel;
		btnCancel.Location = new Point(372, 543);
		btnCancel.Name = AH.A(162103);
		btnCancel.Size = new Size(75, 27);
		btnCancel.TabIndex = 0;
		btnCancel.Text = AH.A(162122);
		btnCancel.UseVisualStyleBackColor = true;
		btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnOk.DialogResult = DialogResult.OK;
		btnOk.Location = new Point(291, 543);
		btnOk.Name = AH.A(162135);
		btnOk.Size = new Size(75, 27);
		btnOk.TabIndex = 1;
		btnOk.Text = AH.A(162146);
		btnOk.UseVisualStyleBackColor = true;
		chkTopic.Appearance = Appearance.Button;
		chkTopic.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		chkTopic.Image = OB.FlysheetStyleTopic;
		chkTopic.ImageAlign = ContentAlignment.TopCenter;
		chkTopic.Location = new Point(15, 24);
		chkTopic.Name = AH.A(162151);
		chkTopic.Padding = new Padding(0, 6, 0, 6);
		chkTopic.Size = new Size(74, 74);
		chkTopic.TabIndex = 2;
		chkTopic.Text = AH.A(3762);
		chkTopic.TextAlign = ContentAlignment.BottomCenter;
		chkTopic.UseVisualStyleBackColor = true;
		chkAgenda.Appearance = Appearance.Button;
		chkAgenda.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		chkAgenda.Image = OB.FlysheetStyleAgenda;
		chkAgenda.ImageAlign = ContentAlignment.TopCenter;
		chkAgenda.Location = new Point(15, 104);
		chkAgenda.Name = AH.A(162168);
		chkAgenda.Padding = new Padding(0, 6, 0, 6);
		chkAgenda.Size = new Size(74, 74);
		chkAgenda.TabIndex = 3;
		chkAgenda.Text = AH.A(122951);
		chkAgenda.TextAlign = ContentAlignment.BottomCenter;
		chkAgenda.UseVisualStyleBackColor = true;
		GroupBox1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		GroupBox1.Controls.Add(RichTextBox2);
		GroupBox1.Controls.Add(RichTextBox1);
		GroupBox1.Controls.Add(chkTopic);
		GroupBox1.Controls.Add(chkAgenda);
		GroupBox1.Font = new Font(AH.A(10041), 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
		GroupBox1.Location = new Point(12, 48);
		GroupBox1.Name = AH.A(162187);
		GroupBox1.Size = new Size(434, 193);
		GroupBox1.TabIndex = 4;
		GroupBox1.TabStop = false;
		GroupBox1.Text = AH.A(162206);
		RichTextBox2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		RichTextBox2.BackColor = SystemColors.Control;
		RichTextBox2.BorderStyle = BorderStyle.None;
		RichTextBox2.Cursor = Cursors.Arrow;
		RichTextBox2.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		RichTextBox2.ForeColor = SystemColors.ControlDarkDark;
		RichTextBox2.Location = new Point(98, 104);
		RichTextBox2.Name = AH.A(162235);
		RichTextBox2.ReadOnly = true;
		RichTextBox2.ScrollBars = RichTextBoxScrollBars.Vertical;
		RichTextBox2.Size = new Size(320, 74);
		RichTextBox2.TabIndex = 9;
		RichTextBox2.Text = AH.A(162260);
		RichTextBox1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		RichTextBox1.BackColor = SystemColors.Control;
		RichTextBox1.BorderStyle = BorderStyle.None;
		RichTextBox1.Cursor = Cursors.Arrow;
		RichTextBox1.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		RichTextBox1.ForeColor = SystemColors.ControlDarkDark;
		RichTextBox1.Location = new Point(98, 24);
		RichTextBox1.Name = AH.A(162608);
		RichTextBox1.ReadOnly = true;
		RichTextBox1.ScrollBars = RichTextBoxScrollBars.Vertical;
		RichTextBox1.Size = new Size(320, 74);
		RichTextBox1.TabIndex = 8;
		RichTextBox1.Text = AH.A(162633);
		grpBehavior.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		grpBehavior.Controls.Add(RichTextBox5);
		grpBehavior.Controls.Add(RichTextBox4);
		grpBehavior.Controls.Add(RichTextBox3);
		grpBehavior.Controls.Add(CheckBox5);
		grpBehavior.Controls.Add(CheckBox4);
		grpBehavior.Controls.Add(CheckBox3);
		grpBehavior.Font = new Font(AH.A(10041), 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
		grpBehavior.Location = new Point(12, 256);
		grpBehavior.Name = AH.A(162895);
		grpBehavior.Size = new Size(434, 274);
		grpBehavior.TabIndex = 5;
		grpBehavior.TabStop = false;
		grpBehavior.Text = AH.A(162918);
		RichTextBox5.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		RichTextBox5.BackColor = SystemColors.Control;
		RichTextBox5.BorderStyle = BorderStyle.None;
		RichTextBox5.Cursor = Cursors.Arrow;
		RichTextBox5.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		RichTextBox5.ForeColor = SystemColors.ControlDarkDark;
		RichTextBox5.Location = new Point(137, 184);
		RichTextBox5.Name = AH.A(162953);
		RichTextBox5.ReadOnly = true;
		RichTextBox5.ScrollBars = RichTextBoxScrollBars.Vertical;
		RichTextBox5.Size = new Size(281, 74);
		RichTextBox5.TabIndex = 12;
		RichTextBox5.Text = AH.A(162978);
		RichTextBox4.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		RichTextBox4.BackColor = SystemColors.Control;
		RichTextBox4.BorderStyle = BorderStyle.None;
		RichTextBox4.Cursor = Cursors.Arrow;
		RichTextBox4.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		RichTextBox4.ForeColor = SystemColors.ControlDarkDark;
		RichTextBox4.Location = new Point(137, 104);
		RichTextBox4.Name = AH.A(163324);
		RichTextBox4.ReadOnly = true;
		RichTextBox4.ScrollBars = RichTextBoxScrollBars.Vertical;
		RichTextBox4.Size = new Size(281, 74);
		RichTextBox4.TabIndex = 11;
		RichTextBox4.Text = AH.A(163349);
		RichTextBox3.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		RichTextBox3.BackColor = SystemColors.Control;
		RichTextBox3.BorderStyle = BorderStyle.None;
		RichTextBox3.Cursor = Cursors.Arrow;
		RichTextBox3.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		RichTextBox3.ForeColor = SystemColors.ControlDarkDark;
		RichTextBox3.Location = new Point(137, 24);
		RichTextBox3.Name = AH.A(163719);
		RichTextBox3.ReadOnly = true;
		RichTextBox3.ScrollBars = RichTextBoxScrollBars.Vertical;
		RichTextBox3.Size = new Size(281, 74);
		RichTextBox3.TabIndex = 10;
		RichTextBox3.Text = AH.A(163744);
		CheckBox5.Appearance = Appearance.Button;
		CheckBox5.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		CheckBox5.Image = OB.SkipDoubleFlysheets;
		CheckBox5.ImageAlign = ContentAlignment.MiddleLeft;
		CheckBox5.Location = new Point(15, 184);
		CheckBox5.Name = AH.A(164092);
		CheckBox5.Padding = new Padding(6, 6, 0, 6);
		CheckBox5.Size = new Size(113, 74);
		CheckBox5.TabIndex = 8;
		CheckBox5.Text = AH.A(164111);
		CheckBox5.UseVisualStyleBackColor = true;
		CheckBox4.Appearance = Appearance.Button;
		CheckBox4.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		CheckBox4.Image = OB.OmitDoubleFlysheets;
		CheckBox4.ImageAlign = ContentAlignment.MiddleLeft;
		CheckBox4.Location = new Point(15, 104);
		CheckBox4.Name = AH.A(164236);
		CheckBox4.Padding = new Padding(6, 6, 0, 6);
		CheckBox4.Size = new Size(113, 74);
		CheckBox4.TabIndex = 7;
		CheckBox4.Text = AH.A(164255);
		CheckBox4.UseVisualStyleBackColor = true;
		CheckBox3.Appearance = Appearance.Button;
		CheckBox3.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		CheckBox3.Image = (Image)componentResourceManager.GetObject(AH.A(164380));
		CheckBox3.ImageAlign = ContentAlignment.MiddleLeft;
		CheckBox3.Location = new Point(15, 24);
		CheckBox3.Name = AH.A(164411);
		CheckBox3.Padding = new Padding(6, 6, 0, 6);
		CheckBox3.Size = new Size(113, 74);
		CheckBox3.TabIndex = 6;
		CheckBox3.Text = AH.A(164430);
		CheckBox3.UseVisualStyleBackColor = true;
		Label6.AutoSize = true;
		Label6.Location = new Point(9, 20);
		Label6.Name = AH.A(164555);
		Label6.Size = new Size(425, 15);
		Label6.TabIndex = 6;
		Label6.Text = AH.A(164568);
		CheckBox6.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
		CheckBox6.AutoSize = true;
		CheckBox6.Location = new Point(12, 549);
		CheckBox6.Name = AH.A(164736);
		CheckBox6.Size = new Size(96, 19);
		CheckBox6.TabIndex = 7;
		CheckBox6.Text = AH.A(164755);
		CheckBox6.UseVisualStyleBackColor = true;
		CheckBox6.Visible = false;
		base.AcceptButton = btnOk;
		base.AutoScaleDimensions = new SizeF(96f, 96f);
		base.AutoScaleMode = AutoScaleMode.Dpi;
		base.CancelButton = btnCancel;
		base.ClientSize = new Size(458, 581);
		base.Controls.Add(CheckBox6);
		base.Controls.Add(Label6);
		base.Controls.Add(grpBehavior);
		base.Controls.Add(GroupBox1);
		base.Controls.Add(btnOk);
		base.Controls.Add(btnCancel);
		DoubleBuffered = true;
		Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		base.FormBorderStyle = FormBorderStyle.FixedDialog;
		base.Icon = (Icon)componentResourceManager.GetObject(AH.A(164784));
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = AH.A(164805);
		base.ShowInTaskbar = false;
		base.StartPosition = FormStartPosition.CenterParent;
		Text = AH.A(164838);
		GroupBox1.ResumeLayout(performLayout: false);
		grpBehavior.ResumeLayout(performLayout: false);
		ResumeLayout(performLayout: false);
		PerformLayout();
	}

	private void A(object A, EventArgs B)
	{
		grpBehavior.Enabled = chkAgenda.Checked;
	}
}
