using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using A;
using ExcelAddIn1.Library2.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class ctpLibrary2 : UserControl
{
	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ElementHost1")]
	private ElementHost m_A;

	internal wpfLibrary A;

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

	public ctpLibrary2()
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
				switch (3)
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
		ElementHost1 = new ElementHost();
		this.A = new wpfLibrary();
		SuspendLayout();
		ElementHost1.BackColor = Color.Transparent;
		ElementHost1.BackColorTransparent = true;
		ElementHost1.Dock = DockStyle.Fill;
		ElementHost1.Font = new Font(VH.A(50021), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		ElementHost1.Location = new Point(0, 0);
		ElementHost1.Name = VH.A(52532);
		ElementHost1.Size = new Size(431, 803);
		ElementHost1.TabIndex = 0;
		ElementHost1.Text = VH.A(52532);
		ElementHost1.Child = this.A;
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		BackColor = Color.White;
		base.Controls.Add(ElementHost1);
		DoubleBuffered = true;
		base.Name = VH.A(169636);
		base.Size = new Size(431, 803);
		ResumeLayout(performLayout: false);
	}
}
