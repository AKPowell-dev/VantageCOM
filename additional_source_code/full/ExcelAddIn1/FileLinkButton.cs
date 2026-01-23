using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelAddIn1;

public sealed class FileLinkButton : Button
{
	private bool A;

	private clsDiscuss.Storage A;

	private string A;

	private Image A;

	public clsDiscuss.Storage Storage
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			Invalidate();
		}
	}

	public override string Text
	{
		get
		{
			if (this.A)
			{
				return this.A;
			}
			return string.Empty;
		}
		set
		{
			this.A = value;
			Invalidate();
		}
	}

	public Image Icon
	{
		get
		{
			if (this.A)
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
						return A;
					}
				}
			}
			return null;
		}
		set
		{
			A = value;
			Invalidate();
		}
	}

	public FileLinkButton()
	{
		this.A = true;
		this.A = clsDiscuss.Storage.Local;
		this.A = string.Empty;
		A = null;
		AutoSize = false;
		base.Margin = new Padding(0, 0, 0, 0);
		TextAlign = ContentAlignment.MiddleLeft;
		base.ImageAlign = ContentAlignment.MiddleLeft;
		base.FlatStyle = FlatStyle.Flat;
		FlatButtonAppearance flatButtonAppearance = base.FlatAppearance;
		flatButtonAppearance.BorderSize = 0;
		flatButtonAppearance.MouseDownBackColor = Color.White;
		flatButtonAppearance.MouseOverBackColor = Color.White;
		_ = null;
		base.UseMnemonic = false;
	}

	protected override void OnPaint(PaintEventArgs e)
	{
		this.A = true;
		Rectangle clientRectangle = base.ClientRectangle;
		TextFormatFlags textFormatFlags = TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter;
		if (!base.UseMnemonic)
		{
			textFormatFlags |= TextFormatFlags.NoPrefix;
		}
		else if (!ShowKeyboardCues)
		{
			textFormatFlags |= TextFormatFlags.HidePrefix;
		}
		checked
		{
			if (!string.IsNullOrEmpty(Text))
			{
				while (true)
				{
					switch (7)
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
				this.A = false;
				base.OnPaint(e);
				this.A = true;
				int num = (int)Math.Round((double)(base.Height - 16) / 2.0);
				int num2 = clientRectangle.Width;
				e.Graphics.DrawImage(Icon, 0, num + 1, 16, 16);
				TextRenderer.DrawText(bounds: new Rectangle(18, 0, num2 - 18, clientRectangle.Height), dc: e.Graphics, text: Text, font: Font, foreColor: ForeColor, flags: textFormatFlags);
			}
			else
			{
				this.A = true;
				base.OnPaint(e);
			}
			clientRectangle = default(Rectangle);
			Rectangle rectangle = default(Rectangle);
			this.A = true;
		}
	}
}
