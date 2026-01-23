using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

[DesignerGenerated]
public sealed class frmImportByName : Form
{
	private IContainer m_A;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("GroupBox1")]
	private GroupBox m_A;

	[AccessedThroughProperty("cbxSourceFile")]
	[CompilerGenerated]
	private ComboBox m_A;

	[AccessedThroughProperty("lblFile")]
	[CompilerGenerated]
	private Label m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("GroupBox2")]
	private GroupBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("lblPaste")]
	private Label m_B;

	[AccessedThroughProperty("txtAddress")]
	[CompilerGenerated]
	private TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxPasteAs")]
	private ComboBox m_B;

	[AccessedThroughProperty("lblAddress")]
	[CompilerGenerated]
	private Label m_C;

	[AccessedThroughProperty("lblName")]
	[CompilerGenerated]
	private Label m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxNames")]
	private ComboBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxFullName")]
	private ComboBox m_D;

	[AccessedThroughProperty("cbxImportType")]
	[CompilerGenerated]
	private ComboBox m_E;

	[AccessedThroughProperty("cbxAddress")]
	[CompilerGenerated]
	private ComboBox m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("PictureBox1")]
	private PictureBox m_A;

	[AccessedThroughProperty("chkViewSource")]
	[CompilerGenerated]
	private CheckBox m_A;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private Microsoft.VisualBasic.Collection m_A;

	private bool m_A;

	internal virtual Button btnOk
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
			EventHandler value2 = E;
			Button button = this.m_A;
			if (button != null)
			{
				while (true)
				{
					switch (2)
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
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EventHandler value2 = F;
			Button button = this.m_B;
			if (button != null)
			{
				while (true)
				{
					switch (5)
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
				button.Click -= value2;
			}
			this.m_B = value;
			button = this.m_B;
			if (button != null)
			{
				button.Click += value2;
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

	internal virtual ComboBox cbxSourceFile
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
			EventHandler value2 = B;
			ComboBox comboBox = this.m_A;
			if (comboBox != null)
			{
				comboBox.SelectedIndexChanged -= value2;
			}
			this.m_A = value;
			comboBox = this.m_A;
			if (comboBox == null)
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				comboBox.SelectedIndexChanged += value2;
				return;
			}
		}
	}

	internal virtual Label lblFile
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

	internal virtual GroupBox GroupBox2
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual Label lblPaste
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual TextBox txtAddress
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

	internal virtual ComboBox cbxPasteAs
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual Label lblAddress
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual Label lblName
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual ComboBox cbxNames
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EventHandler value2 = C;
			ComboBox comboBox = this.m_C;
			if (comboBox != null)
			{
				while (true)
				{
					switch (2)
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
				comboBox.SelectedIndexChanged -= value2;
			}
			this.m_C = value;
			comboBox = this.m_C;
			if (comboBox != null)
			{
				comboBox.SelectedIndexChanged += value2;
			}
		}
	}

	internal virtual ComboBox cbxFullName
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual ComboBox cbxImportType
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal virtual ComboBox cbxAddress
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal virtual PictureBox PictureBox1
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

	internal virtual CheckBox chkViewSource
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
			EventHandler value2 = D;
			CheckBox checkBox = this.m_A;
			if (checkBox != null)
			{
				while (true)
				{
					switch (5)
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
			if (checkBox == null)
			{
				return;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				checkBox.CheckedChanged += value2;
				return;
			}
		}
	}

	public frmImportByName()
	{
		base.Load += A;
		base.FormClosing += A;
		this.m_A = true;
		System.Windows.Forms.Application.EnableVisualStyles();
		A();
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (!disposing)
			{
				return;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (this.m_A == null)
				{
					return;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					this.m_A.Dispose();
					return;
				}
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
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmImportByName));
		btnOk = new Button();
		btnCancel = new Button();
		GroupBox1 = new GroupBox();
		chkViewSource = new CheckBox();
		cbxSourceFile = new ComboBox();
		lblFile = new Label();
		txtAddress = new TextBox();
		lblAddress = new Label();
		cbxNames = new ComboBox();
		lblName = new Label();
		GroupBox2 = new GroupBox();
		lblPaste = new Label();
		cbxPasteAs = new ComboBox();
		cbxFullName = new ComboBox();
		cbxImportType = new ComboBox();
		cbxAddress = new ComboBox();
		PictureBox1 = new PictureBox();
		GroupBox1.SuspendLayout();
		GroupBox2.SuspendLayout();
		((ISupportInitialize)PictureBox1).BeginInit();
		SuspendLayout();
		btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnOk.Enabled = false;
		btnOk.FlatStyle = FlatStyle.System;
		btnOk.Location = new Point(254, 277);
		btnOk.Name = AH.A(162135);
		btnOk.Size = new Size(78, 27);
		btnOk.TabIndex = 0;
		btnOk.Text = AH.A(164867);
		btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnCancel.DialogResult = DialogResult.Cancel;
		btnCancel.FlatStyle = FlatStyle.System;
		btnCancel.Location = new Point(338, 277);
		btnCancel.Name = AH.A(162103);
		btnCancel.Size = new Size(78, 27);
		btnCancel.TabIndex = 1;
		btnCancel.Text = AH.A(164874);
		GroupBox1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		GroupBox1.Controls.Add(chkViewSource);
		GroupBox1.Controls.Add(cbxSourceFile);
		GroupBox1.Controls.Add(lblFile);
		GroupBox1.Controls.Add(txtAddress);
		GroupBox1.Controls.Add(lblAddress);
		GroupBox1.Controls.Add(cbxNames);
		GroupBox1.Controls.Add(lblName);
		GroupBox1.FlatStyle = FlatStyle.System;
		GroupBox1.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
		GroupBox1.Location = new Point(12, 18);
		GroupBox1.Name = AH.A(162187);
		GroupBox1.Size = new Size(404, 157);
		GroupBox1.TabIndex = 1;
		GroupBox1.TabStop = false;
		GroupBox1.Text = AH.A(96758);
		chkViewSource.Appearance = Appearance.Button;
		chkViewSource.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		chkViewSource.Image = OB.Find;
		chkViewSource.ImageAlign = ContentAlignment.MiddleLeft;
		chkViewSource.Location = new Point(97, 113);
		chkViewSource.Name = AH.A(164889);
		chkViewSource.Padding = new Padding(6, 0, 0, 0);
		chkViewSource.Size = new Size(113, 27);
		chkViewSource.TabIndex = 9;
		chkViewSource.Text = AH.A(164916);
		chkViewSource.TextAlign = ContentAlignment.MiddleCenter;
		chkViewSource.UseVisualStyleBackColor = true;
		cbxSourceFile.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		cbxSourceFile.DropDownStyle = ComboBoxStyle.DropDownList;
		cbxSourceFile.FlatStyle = FlatStyle.System;
		cbxSourceFile.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		cbxSourceFile.FormattingEnabled = true;
		cbxSourceFile.Location = new Point(98, 26);
		cbxSourceFile.Name = AH.A(164949);
		cbxSourceFile.Size = new Size(289, 23);
		cbxSourceFile.TabIndex = 3;
		lblFile.AutoSize = true;
		lblFile.FlatStyle = FlatStyle.System;
		lblFile.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		lblFile.Location = new Point(18, 29);
		lblFile.Name = AH.A(164976);
		lblFile.Size = new Size(64, 15);
		lblFile.TabIndex = 2;
		lblFile.Text = AH.A(164991);
		lblFile.TextAlign = ContentAlignment.MiddleLeft;
		txtAddress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		txtAddress.BackColor = Color.White;
		txtAddress.Enabled = false;
		txtAddress.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		txtAddress.ForeColor = SystemColors.ControlDark;
		txtAddress.Location = new Point(98, 84);
		txtAddress.Name = AH.A(165016);
		txtAddress.ReadOnly = true;
		txtAddress.Size = new Size(289, 23);
		txtAddress.TabIndex = 4;
		lblAddress.AutoSize = true;
		lblAddress.FlatStyle = FlatStyle.System;
		lblAddress.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		lblAddress.Location = new Point(18, 87);
		lblAddress.Name = AH.A(165037);
		lblAddress.Size = new Size(49, 15);
		lblAddress.TabIndex = 2;
		lblAddress.Text = AH.A(165058);
		lblAddress.TextAlign = ContentAlignment.MiddleLeft;
		cbxNames.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		cbxNames.DropDownStyle = ComboBoxStyle.DropDownList;
		cbxNames.FlatStyle = FlatStyle.System;
		cbxNames.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		cbxNames.FormattingEnabled = true;
		cbxNames.Location = new Point(98, 55);
		cbxNames.Name = AH.A(165075);
		cbxNames.Size = new Size(289, 23);
		cbxNames.TabIndex = 0;
		lblName.AutoSize = true;
		lblName.FlatStyle = FlatStyle.System;
		lblName.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		lblName.Location = new Point(18, 58);
		lblName.Name = AH.A(165092);
		lblName.Size = new Size(75, 15);
		lblName.TabIndex = 1;
		lblName.Text = AH.A(165107);
		lblName.TextAlign = ContentAlignment.MiddleLeft;
		GroupBox2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		GroupBox2.Controls.Add(lblPaste);
		GroupBox2.Controls.Add(cbxPasteAs);
		GroupBox2.FlatStyle = FlatStyle.System;
		GroupBox2.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
		GroupBox2.Location = new Point(12, 190);
		GroupBox2.Name = AH.A(165130);
		GroupBox2.Size = new Size(404, 68);
		GroupBox2.TabIndex = 6;
		GroupBox2.TabStop = false;
		GroupBox2.Text = AH.A(165149);
		lblPaste.AutoSize = true;
		lblPaste.FlatStyle = FlatStyle.System;
		lblPaste.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		lblPaste.Location = new Point(18, 29);
		lblPaste.Name = AH.A(165164);
		lblPaste.Size = new Size(59, 15);
		lblPaste.TabIndex = 5;
		lblPaste.Text = AH.A(165181);
		lblPaste.TextAlign = ContentAlignment.MiddleLeft;
		cbxPasteAs.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		cbxPasteAs.DropDownStyle = ComboBoxStyle.DropDownList;
		cbxPasteAs.FlatStyle = FlatStyle.System;
		cbxPasteAs.Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		cbxPasteAs.FormattingEnabled = true;
		cbxPasteAs.Items.AddRange(new object[5]
		{
			AH.A(3293),
			AH.A(165202),
			AH.A(165235),
			AH.A(165272),
			AH.A(70464)
		});
		cbxPasteAs.Location = new Point(98, 26);
		cbxPasteAs.Name = AH.A(165283);
		cbxPasteAs.Size = new Size(289, 23);
		cbxPasteAs.TabIndex = 3;
		cbxFullName.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		cbxFullName.DropDownStyle = ComboBoxStyle.DropDownList;
		cbxFullName.FlatStyle = FlatStyle.System;
		cbxFullName.FormattingEnabled = true;
		cbxFullName.Location = new Point(157, 278);
		cbxFullName.Name = AH.A(165304);
		cbxFullName.Size = new Size(38, 23);
		cbxFullName.TabIndex = 8;
		cbxFullName.Visible = false;
		cbxImportType.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		cbxImportType.DropDownStyle = ComboBoxStyle.DropDownList;
		cbxImportType.FlatStyle = FlatStyle.System;
		cbxImportType.FormattingEnabled = true;
		cbxImportType.Items.AddRange(new object[5]
		{
			AH.A(9078),
			AH.A(9081),
			AH.A(9084),
			AH.A(9090),
			AH.A(9087)
		});
		cbxImportType.Location = new Point(203, 278);
		cbxImportType.Name = AH.A(165327);
		cbxImportType.Size = new Size(39, 23);
		cbxImportType.TabIndex = 9;
		cbxImportType.Visible = false;
		cbxAddress.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		cbxAddress.DropDownStyle = ComboBoxStyle.DropDownList;
		cbxAddress.FlatStyle = FlatStyle.System;
		cbxAddress.FormattingEnabled = true;
		cbxAddress.Location = new Point(112, 278);
		cbxAddress.Name = AH.A(165354);
		cbxAddress.Size = new Size(38, 23);
		cbxAddress.TabIndex = 10;
		cbxAddress.Visible = false;
		PictureBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
		PictureBox1.BackColor = SystemColors.Control;
		PictureBox1.Image = (Image)componentResourceManager.GetObject(AH.A(165375));
		PictureBox1.Location = new Point(12, 287);
		PictureBox1.Name = AH.A(165410);
		PictureBox1.Size = new Size(57, 14);
		PictureBox1.TabIndex = 94;
		PictureBox1.TabStop = false;
		base.AcceptButton = btnOk;
		base.AutoScaleDimensions = new SizeF(96f, 96f);
		base.AutoScaleMode = AutoScaleMode.Dpi;
		base.CancelButton = btnCancel;
		base.ClientSize = new Size(428, 316);
		base.Controls.Add(PictureBox1);
		base.Controls.Add(btnCancel);
		base.Controls.Add(btnOk);
		base.Controls.Add(cbxAddress);
		base.Controls.Add(cbxFullName);
		base.Controls.Add(cbxImportType);
		base.Controls.Add(GroupBox2);
		base.Controls.Add(GroupBox1);
		Font = new System.Drawing.Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		base.FormBorderStyle = FormBorderStyle.FixedDialog;
		base.Icon = (Icon)componentResourceManager.GetObject(AH.A(164784));
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = AH.A(165433);
		base.ShowInTaskbar = false;
		base.StartPosition = FormStartPosition.CenterParent;
		Text = AH.A(165464);
		base.TopMost = true;
		GroupBox1.ResumeLayout(performLayout: false);
		GroupBox1.PerformLayout();
		GroupBox2.ResumeLayout(performLayout: false);
		GroupBox2.PerformLayout();
		((ISupportInitialize)PictureBox1).EndInit();
		ResumeLayout(performLayout: false);
	}

	private void A(object A, EventArgs B)
	{
		int num = 0;
		try
		{
			this.m_A = (Microsoft.Office.Interop.Excel.Application)Interaction.GetObject(null, AH.A(93301));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			num = this.m_A.Workbooks.Count;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		if (num == 0)
		{
			Close();
			this.B(AH.A(165493));
			return;
		}
		this.m_A = NG.A.Application;
		this.m_A = this.m_A.ActivePresentation;
		foreach (Workbook workbook2 in this.m_A.Workbooks)
		{
			if (Strings.Len(workbook2.Path) > 0)
			{
				cbxSourceFile.Items.Add(workbook2.Name);
				cbxFullName.Items.Add(workbook2.FullName);
			}
			Workbook workbook = null;
		}
		this.m_A = false;
		cbxPasteAs.SelectedIndex = 0;
		cbxSourceFile.SelectedIndex = 0;
	}

	private void A(object A, FormClosingEventArgs B)
	{
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
	}

	private void B(object A, EventArgs B)
	{
		if (this.m_A)
		{
			return;
		}
		this.m_A = true;
		this.m_A = new Microsoft.VisualBasic.Collection();
		string index = Conversions.ToString(cbxSourceFile.SelectedItem);
		cbxNames.Items.Clear();
		cbxAddress.Items.Clear();
		txtAddress.Text = "";
		btnOk.Enabled = false;
		chkViewSource.Enabled = false;
		try
		{
			Workbook workbook = this.m_A.Workbooks[index];
			try
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = workbook.Names.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Name name = (Name)enumerator.Current;
						Name name2 = name;
						if (name2.Visible && (name2.Name.StartsWith(AH.A(165550)) | KG.A.ShowAllNames))
						{
							while (true)
							{
								switch (5)
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
							if ((!this.A(name2.Name) & !this.A(name)) && name2.RefersToRange.Cells.Count > 1)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									break;
								}
								cbxNames.Items.Add(name.Name);
								cbxAddress.Items.Add(Ranges.RangeAddress(name.RefersToRange));
								this.m_A.Add(name.RefersToRange);
							}
						}
						name2 = null;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_01b1;
						}
						continue;
						end_IL_01b1:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				workbook = null;
				this.B();
				this.m_A = false;
				if (cbxNames.Items.Count > 0)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						cbxNames.SelectedIndex = 0;
						break;
					}
				}
				else
				{
					this.B(AH.A(165553));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			this.A(AH.A(165709));
			ProjectData.ClearProjectError();
		}
		this.m_A = false;
	}

	private bool A(Name A)
	{
		return LikeOperator.LikeString(A.RefersTo.ToString(), AH.A(165828), CompareMethod.Binary) | LikeOperator.LikeString(A.RefersTo.ToString(), AH.A(165841), CompareMethod.Binary);
	}

	private bool A(string A)
	{
		return LikeOperator.LikeString(A, AH.A(165856), CompareMethod.Binary) | LikeOperator.LikeString(A, AH.A(165879), CompareMethod.Binary);
	}

	private void C(object A, EventArgs B)
	{
		if (this.m_A)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			string text = Conversions.ToString(cbxAddress.Items[cbxNames.SelectedIndex]);
			txtAddress.Text = text;
			this.B();
			this.m_A = false;
			return;
		}
	}

	private void D(object A, EventArgs B)
	{
		if (this.m_A)
		{
			return;
		}
		if (chkViewSource.Checked)
		{
			if (Operators.CompareString(txtAddress.Text, "", TextCompare: false) == 0)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				try
				{
					C();
					Interaction.AppActivate(this.m_A.Caption);
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					this.A(AH.A(165906));
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
		Interaction.AppActivate(this.m_A.Caption);
		this.m_A.Windows[1].Activate();
	}

	private void E(object A, EventArgs B)
	{
		if (Operators.CompareString(txtAddress.Text, "", TextCompare: false) == 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.B(AH.A(165989));
					return;
				}
			}
		}
		try
		{
			C();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			this.A(AH.A(166044));
			ProjectData.ClearProjectError();
		}
		try
		{
			Interaction.AppActivate(this.m_A.Caption);
			BringToFront();
			base.DialogResult = DialogResult.OK;
			Close();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	private void F(object A, EventArgs B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 88:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0021;
						case 4:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 5:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0021:
					num2 = 3;
					base.DialogResult = DialogResult.Cancel;
					break;
					IL_0007:
					num2 = 2;
					this.m_A.Windows[1].Activate();
					goto IL_0021;
					end_IL_0000_2:
					break;
				}
				num2 = 4;
				Close();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 88;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void A(bool A)
	{
		lblFile.Enabled = A;
		lblName.Enabled = A;
		cbxSourceFile.Enabled = A;
		cbxNames.Enabled = A;
	}

	private void B(bool A)
	{
		lblAddress.Enabled = !A;
		TextBox textBox = txtAddress;
		textBox.ReadOnly = A;
		if (A)
		{
			while (true)
			{
				switch (6)
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
			textBox.ForeColor = Color.DarkGray;
		}
		else
		{
			textBox.ForeColor = Color.Black;
		}
		textBox = null;
	}

	private void B()
	{
		if (Strings.Len(txtAddress.Text) == 0)
		{
			btnOk.Enabled = false;
			chkViewSource.Enabled = false;
		}
		else
		{
			btnOk.Enabled = true;
			chkViewSource.Enabled = true;
		}
	}

	private void C()
	{
		Range obj = (Range)this.m_A[checked(cbxNames.SelectedIndex + 1)];
		((Workbook)obj.Worksheet.Parent).Activate();
		obj.Worksheet.Activate();
		obj.Select();
		_ = null;
	}

	private void A(string A)
	{
		MessageBox.Show(A, AH.A(5874), MessageBoxButtons.OK, MessageBoxIcon.Hand);
	}

	private void B(string A)
	{
		MessageBox.Show(A, AH.A(5874), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	}
}
