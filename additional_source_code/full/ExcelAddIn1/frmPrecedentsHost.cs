using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using A;
using ExcelAddIn1.Audit;
using ExcelAddIn1.Audit.TraceDialogs;
using ExcelAddIn1.Audit.TraceDialogs.Precedents;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class frmPrecedentsHost : TraceForm
{
	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ElementHost1")]
	private ElementHost m_A;

	private const int m_A = 28;

	[CompilerGenerated]
	private bool m_A;

	internal virtual ElementHost ElementHost1
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

	public bool IgnoreDeactivate
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public frmPrecedentsHost()
	{
		base.Load += A;
		base.FormClosing += A;
		A();
		Base.SetFormSizeAndPosition(this);
		SetPadding();
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (disposing && this.m_A != null)
			{
				this.m_A.Dispose();
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
		ElementHost1 = new ElementHost();
		SuspendLayout();
		ElementHost1.BackColor = Color.White;
		ElementHost1.Dock = DockStyle.Fill;
		ElementHost1.Font = new Font(VH.A(50021), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		ElementHost1.Location = new Point(2, 30);
		ElementHost1.Name = VH.A(52532);
		ElementHost1.Size = new Size(424, 277);
		ElementHost1.TabIndex = 0;
		ElementHost1.TabStop = false;
		ElementHost1.Text = VH.A(52532);
		ElementHost1.Visible = false;
		ElementHost1.Child = null;
		base.AutoScaleDimensions = new SizeF(7f, 15f);
		base.AutoScaleMode = AutoScaleMode.Font;
		AutoValidate = AutoValidate.EnablePreventFocusChange;
		base.ClientSize = new Size(428, 309);
		base.Controls.Add(ElementHost1);
		MinimumSize = new Size(350, 250);
		base.Name = VH.A(52638);
		base.Padding = new Padding(2, 30, 2, 2);
		Text = VH.A(52673);
		ResumeLayout(performLayout: false);
	}

	protected override void WndProc(ref Message m)
	{
		if (m.Msg == 28)
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
			if (m.WParam.ToInt32() == 0)
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
				if (!IgnoreDeactivate)
				{
					((wpfPrecedents)ElementHost1.Child).Close();
				}
			}
		}
		base.WndProc(ref m);
	}

	private void A(object A, EventArgs B)
	{
		try
		{
			ElementHost1.Child = new wpfPrecedents(this);
			ElementHost1.Visible = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		base.Activated += FormActivated;
		base.Deactivate += FormDeactivate;
	}

	private void A(object A, FormClosingEventArgs B)
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
				case 132:
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
							goto IL_000f;
						case 4:
							goto IL_001f;
						case 5:
							goto IL_002e;
						case 6:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 7:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_002e:
					num2 = 5;
					base.Activated -= FormActivated;
					break;
					IL_0007:
					num2 = 2;
					Base.SaveFormSizeAndPosition(this);
					goto IL_000f;
					IL_000f:
					num2 = 3;
					ElementHost1.Child = null;
					goto IL_001f;
					IL_001f:
					num2 = 4;
					ElementHost1.Dispose();
					goto IL_002e;
					end_IL_0000_2:
					break;
				}
				num2 = 6;
				base.Deactivate -= FormDeactivate;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 132;
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
			switch (6)
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

	public void FormActivated(object sender, EventArgs e)
	{
		base.Opacity = 1.0;
	}

	public void FormDeactivate(object sender, EventArgs e)
	{
		if (IgnoreDeactivate)
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
			base.Opacity = 0.7;
			if (!K.Settings.AuditHighlightCells)
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
				Base.RemoveHighlight();
				return;
			}
		}
	}
}
