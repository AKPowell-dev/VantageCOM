using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using A;
using ExcelAddIn1.Charts.MoveDataLabels;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class ctpMoveDataLabels : UserControl
{
	private IContainer m_A;

	[AccessedThroughProperty("ElementHost1")]
	[CompilerGenerated]
	private ElementHost m_A;

	internal wpfPane A;

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

	public ctpMoveDataLabels()
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
				switch (4)
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
		ElementHost1 = new ElementHost();
		this.A = new wpfPane();
		SuspendLayout();
		ElementHost1.BackColor = Color.Transparent;
		ElementHost1.Dock = DockStyle.Fill;
		ElementHost1.Font = new Font(VH.A(50021), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		ElementHost1.Location = new Point(0, 0);
		ElementHost1.Name = VH.A(52532);
		ElementHost1.Size = new Size(497, 336);
		ElementHost1.TabIndex = 0;
		ElementHost1.Text = VH.A(52532);
		ElementHost1.Child = this.A;
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		BackColor = Color.White;
		base.Controls.Add(ElementHost1);
		DoubleBuffered = true;
		base.Name = VH.A(82374);
		base.Size = new Size(497, 336);
		ResumeLayout(performLayout: false);
	}
}
