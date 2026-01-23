using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using A;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.FormatPainter;

namespace PowerPointAddIn1;

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
			if (!disposing || this.m_A == null)
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
		wpfElementHost = new ElementHost();
		this.A = new FormatTree();
		SuspendLayout();
		wpfElementHost.BackColor = Color.Transparent;
		wpfElementHost.BackColorTransparent = true;
		wpfElementHost.Dock = DockStyle.Fill;
		wpfElementHost.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		wpfElementHost.Location = new Point(0, 0);
		wpfElementHost.Name = AH.A(152019);
		wpfElementHost.Size = new Size(419, 620);
		wpfElementHost.TabIndex = 0;
		wpfElementHost.Text = AH.A(10058);
		wpfElementHost.Child = this.A;
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		BackColor = Color.Transparent;
		base.Controls.Add(wpfElementHost);
		base.Name = AH.A(152048);
		base.Size = new Size(419, 620);
		ResumeLayout(performLayout: false);
	}
}
