using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using A;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.UI;

namespace PowerPointAddIn1;

[DesignerGenerated]
public sealed class ctpDeckCheck : UserControl
{
	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ElementHost1")]
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

	public ctpDeckCheck(bool blnWholePresentation)
	{
		A();
		((wpfPane)ElementHost1.Child).WholePresentation = blnWholePresentation;
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
				switch (7)
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
		this.A = new wpfPane();
		SuspendLayout();
		ElementHost1.BackColor = Color.Transparent;
		ElementHost1.BackColorTransparent = true;
		ElementHost1.Dock = DockStyle.Fill;
		ElementHost1.Font = new Font(AH.A(10041), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		ElementHost1.Location = new Point(0, 0);
		ElementHost1.Name = AH.A(10058);
		ElementHost1.Size = new Size(361, 635);
		ElementHost1.TabIndex = 0;
		ElementHost1.Text = AH.A(10058);
		ElementHost1.Child = this.A;
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		base.Controls.Add(ElementHost1);
		base.Name = AH.A(118754);
		base.Size = new Size(361, 635);
		ResumeLayout(performLayout: false);
	}
}
