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
public sealed class frmSectionRename : Form
{
	private IContainer m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button m_A;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("txtName")]
	private TextBox m_A;

	[AccessedThroughProperty("Label1")]
	[CompilerGenerated]
	private Label m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ToolTip1")]
	private ToolTip m_A;

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
			EventHandler value2 = B;
			Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnOk
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
			EventHandler value2 = A;
			Button button = this.m_B;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual TextBox txtName
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

	internal virtual Label Label1
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

	internal virtual ToolTip ToolTip1
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

	public frmSectionRename()
	{
		Application.EnableVisualStyles();
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
				switch (1)
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
					switch (4)
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
		this.m_A = new Container();
		btnCancel = new Button();
		btnOk = new Button();
		txtName = new TextBox();
		Label1 = new Label();
		ToolTip1 = new ToolTip(this.m_A);
		SuspendLayout();
		btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnCancel.DialogResult = DialogResult.Cancel;
		btnCancel.Location = new Point(141, 57);
		btnCancel.Name = AH.A(162103);
		btnCancel.Size = new Size(75, 25);
		btnCancel.TabIndex = 0;
		btnCancel.Text = AH.A(162122);
		btnCancel.UseVisualStyleBackColor = true;
		btnOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnOk.DialogResult = DialogResult.Cancel;
		btnOk.Location = new Point(60, 57);
		btnOk.Name = AH.A(162135);
		btnOk.Size = new Size(75, 25);
		btnOk.TabIndex = 1;
		btnOk.Text = AH.A(166143);
		btnOk.UseVisualStyleBackColor = true;
		txtName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		txtName.Location = new Point(10, 27);
		txtName.Name = AH.A(166158);
		txtName.Size = new Size(205, 23);
		txtName.TabIndex = 2;
		ToolTip1.SetToolTip(txtName, AH.A(166173));
		Label1.AutoSize = true;
		Label1.Location = new Point(7, 6);
		Label1.Name = AH.A(166258);
		Label1.Size = new Size(82, 15);
		Label1.TabIndex = 3;
		Label1.Text = AH.A(166271);
		base.AcceptButton = btnOk;
		base.AutoScaleDimensions = new SizeF(96f, 96f);
		base.AutoScaleMode = AutoScaleMode.Dpi;
		base.CancelButton = btnCancel;
		base.ClientSize = new Size(224, 90);
		base.Controls.Add(Label1);
		base.Controls.Add(txtName);
		base.Controls.Add(btnOk);
		base.Controls.Add(btnCancel);
		Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		base.FormBorderStyle = FormBorderStyle.FixedDialog;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = AH.A(166298);
		base.ShowIcon = false;
		base.ShowInTaskbar = false;
		base.StartPosition = FormStartPosition.CenterScreen;
		Text = AH.A(166331);
		ResumeLayout(performLayout: false);
		PerformLayout();
	}

	private void A(object A, EventArgs B)
	{
		base.DialogResult = DialogResult.OK;
		Close();
	}

	private void B(object A, EventArgs B)
	{
		base.DialogResult = DialogResult.Cancel;
		Close();
	}
}
