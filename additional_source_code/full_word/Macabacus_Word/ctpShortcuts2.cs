using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using A;
using Macabacus_Word.Keyboard;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word;

[DesignerGenerated]
public sealed class ctpShortcuts2 : UserControl
{
	private IContainer m_A;

	internal wpfShortcuts A;

	[AccessedThroughProperty("ElementHost1")]
	[CompilerGenerated]
	private ElementHost m_A;

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

	public ctpShortcuts2()
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
					switch (1)
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
		ElementHost1 = new ElementHost();
		this.A = new wpfShortcuts();
		SuspendLayout();
		ElementHost1.BackColor = Color.Transparent;
		ElementHost1.BackColorTransparent = true;
		ElementHost1.Dock = DockStyle.Fill;
		ElementHost1.Font = new Font(XC.A(1017), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		ElementHost1.Location = new Point(0, 0);
		ElementHost1.Name = XC.A(1034);
		ElementHost1.Size = new Size(370, 950);
		ElementHost1.TabIndex = 0;
		ElementHost1.Text = XC.A(1034);
		ElementHost1.Child = this.A;
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		base.Controls.Add(ElementHost1);
		base.Name = XC.A(22298);
		base.Size = new Size(370, 950);
		ResumeLayout(performLayout: false);
	}
}
