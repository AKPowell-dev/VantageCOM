using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class frmExplorerPreview : Form
{
	private IContainer m_A;

	[AccessedThroughProperty("btnClose")]
	[CompilerGenerated]
	private Button m_A;

	private const int m_A = 28;

	public bool DrawBorder;

	internal virtual Button btnClose
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

	public frmExplorerPreview()
	{
		DrawBorder = true;
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
				switch (3)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (this.m_A != null)
				{
					this.m_A.Dispose();
				}
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
		btnClose = new Button();
		SuspendLayout();
		btnClose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
		btnClose.DialogResult = DialogResult.Cancel;
		btnClose.FlatAppearance.BorderSize = 0;
		btnClose.FlatAppearance.MouseDownBackColor = Color.FromArgb(237, 237, 237);
		btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(237, 237, 237);
		btnClose.FlatStyle = FlatStyle.Flat;
		btnClose.Location = new Point(609, 447);
		btnClose.Name = VH.A(198558);
		btnClose.Size = new Size(17, 17);
		btnClose.TabIndex = 1;
		btnClose.TabStop = false;
		btnClose.UseVisualStyleBackColor = true;
		btnClose.Visible = false;
		base.AutoScaleDimensions = new SizeF(96f, 96f);
		base.AutoScaleMode = AutoScaleMode.Dpi;
		BackColor = Color.White;
		BackgroundImageLayout = ImageLayout.Center;
		base.CancelButton = btnClose;
		base.ClientSize = new Size(628, 549);
		base.Controls.Add(btnClose);
		DoubleBuffered = true;
		Font = new Font(VH.A(50021), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		base.FormBorderStyle = FormBorderStyle.None;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.Name = VH.A(198575);
		base.ShowInTaskbar = false;
		base.StartPosition = FormStartPosition.Manual;
		Text = VH.A(198612);
		ResumeLayout(performLayout: false);
	}

	protected override void WndProc(ref Message m)
	{
		if (m.Msg == 28)
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
			if (m.WParam.ToInt32() == 0)
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
				Hide();
			}
		}
		base.WndProc(ref m);
	}

	protected override void OnPaint(PaintEventArgs e)
	{
		base.OnPaint(e);
		if (!DrawBorder)
		{
			return;
		}
		checked
		{
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
				Point point = new Point(0, 0);
				Point point2 = new Point(base.Width - 1, 0);
				Point point3 = new Point(base.Width - 1, base.Height - 1);
				Point point4 = new Point(0, base.Height - 1);
				Pen pen = new Pen(Color.FromArgb(198, 198, 198), 1f);
				Graphics graphics = e.Graphics;
				graphics.DrawLine(pen, point, point2);
				graphics.DrawLine(pen, point2, point3);
				graphics.DrawLine(pen, point3, point4);
				graphics.DrawLine(pen, point4, point);
				_ = null;
				pen.Dispose();
				point = default(Point);
				point2 = default(Point);
				point3 = default(Point);
				point4 = default(Point);
				return;
			}
		}
	}
}
