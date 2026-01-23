using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using A;
using ExcelAddIn1.FormatPainter;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class ctpFormatPainter : UserControl
{
	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("wpfElementHost")]
	private ElementHost m_A;

	internal FormatTree A;

	internal virtual ElementHost wpfElementHost
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

	public ctpFormatPainter()
	{
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
		wpfElementHost = new ElementHost();
		this.A = new FormatTree();
		SuspendLayout();
		wpfElementHost.BackColor = Color.White;
		wpfElementHost.Dock = DockStyle.Fill;
		wpfElementHost.Font = new Font(VH.A(50021), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		wpfElementHost.Location = new Point(0, 0);
		wpfElementHost.Name = VH.A(173662);
		wpfElementHost.Size = new Size(419, 620);
		wpfElementHost.TabIndex = 0;
		wpfElementHost.Text = VH.A(52532);
		wpfElementHost.Child = this.A;
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		BackColor = Color.White;
		base.Controls.Add(wpfElementHost);
		base.Name = VH.A(173691);
		base.Size = new Size(419, 620);
		ResumeLayout(performLayout: false);
	}
}
