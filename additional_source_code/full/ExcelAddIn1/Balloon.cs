using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelAddIn1;

public sealed class Balloon : Panel
{
	public enum BalloonStyleEnum
	{
		CommentFromMe = 1,
		AttachmentFromMe,
		CommentFromOthers,
		AttachmentFromOthers
	}

	public enum BalloonContentEnum
	{
		Text = 1,
		File,
		Link,
		ScreenShot
	}

	private BalloonContentEnum m_A;

	private bool m_A;

	private int m_A;

	public BalloonContentEnum BalloonContent
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public bool AuthorIsMe
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			if (value)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						base.Margin = new Padding(0, 6, 20, 0);
						ForeColor = Color.White;
						return;
					}
				}
			}
			base.Margin = new Padding(20, 6, 0, 0);
			ForeColor = Color.FromKnownColor(KnownColor.ControlText);
		}
	}

	public int Radius
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public Balloon()
	{
		this.m_A = BalloonContentEnum.Text;
		AuthorIsMe = true;
		AutoSize = true;
		base.BorderStyle = BorderStyle.None;
		base.Margin = new Padding(0, 6, 20, 0);
		base.Padding = new Padding(0, 0, 0, 0);
		ForeColor = Color.White;
		base.TabStop = false;
		Radius = 7;
	}

	protected override void OnPaint(PaintEventArgs e)
	{
		Color color = Color.FromKnownColor(KnownColor.MenuHighlight);
		Color gainsboro = Color.Gainsboro;
		Rectangle clientRectangle = base.ClientRectangle;
		Rectangle a = checked(new Rectangle(clientRectangle.X, clientRectangle.Y, clientRectangle.Width - 1, clientRectangle.Height - 1));
		SmoothingMode smoothingMode = e.Graphics.SmoothingMode;
		GraphicsPath graphicsPath = A(a, Radius);
		Pen pen;
		Brush brush;
		if (AuthorIsMe)
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
			pen = new Pen(color, 1f);
			brush = ((BalloonContent != BalloonContentEnum.Text) ? new SolidBrush(Color.White) : new SolidBrush(color));
		}
		else
		{
			pen = new Pen(new SolidBrush(gainsboro), 1f);
			if (BalloonContent == BalloonContentEnum.Text)
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
				brush = new SolidBrush(gainsboro);
			}
			else
			{
				brush = new SolidBrush(Color.White);
			}
		}
		e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
		A(e.Graphics, graphicsPath, brush);
		A(e.Graphics, graphicsPath, pen);
		e.Graphics.SmoothingMode = smoothingMode;
		base.OnPaint(e);
		brush.Dispose();
		pen.Dispose();
		graphicsPath.Dispose();
	}

	private void A(Graphics A, GraphicsPath B, Brush C)
	{
		A.FillPath(C, B);
	}

	private void A(Graphics A, GraphicsPath B, Pen C)
	{
		A.DrawPath(C, B);
	}

	private GraphicsPath A(Rectangle A, int B)
	{
		GraphicsPath graphicsPath = new GraphicsPath();
		checked
		{
			int num = B * 2;
			GraphicsPath graphicsPath2 = graphicsPath;
			if (AuthorIsMe)
			{
				graphicsPath2.AddLine(A.Left + num + 6, A.Top, A.Right - num, A.Top);
				graphicsPath2.AddArc(Rectangle.FromLTRB(A.Right - num, A.Top, A.Right, A.Top + num), -90f, 90f);
				graphicsPath2.AddLine(A.Right, A.Top + num, A.Right, A.Bottom - num);
				graphicsPath2.AddArc(Rectangle.FromLTRB(A.Right - num, A.Bottom - num, A.Right, A.Bottom), 0f, 90f);
				graphicsPath2.AddLine(A.Right - num, A.Bottom, A.Left, A.Bottom);
				graphicsPath2.AddLine(A.Left, A.Bottom, A.Left + 6, A.Bottom - 8);
				graphicsPath2.AddArc(Rectangle.FromLTRB(A.Left + 6, A.Top, A.Left + num + 6, A.Top + num), 180f, 90f);
			}
			else
			{
				graphicsPath2.AddLine(A.Left + num, A.Top, A.Right - num - 6, A.Top);
				graphicsPath2.AddArc(Rectangle.FromLTRB(A.Right - num - 6, A.Top, A.Right - 6, A.Top + num), -90f, 90f);
				graphicsPath2.AddLine(A.Right - 6, A.Top + num, A.Right - 6, A.Bottom - 8);
				graphicsPath2.AddLine(A.Right - 6, A.Bottom - 8, A.Right, A.Bottom);
				graphicsPath2.AddLine(A.Right, A.Bottom, A.Left + num, A.Bottom);
				graphicsPath2.AddArc(Rectangle.FromLTRB(A.Left, A.Bottom - num, A.Left + num, A.Bottom), 90f, 90f);
				graphicsPath2.AddLine(A.Left, A.Bottom - num, A.Left, A.Top + num);
				graphicsPath2.AddArc(Rectangle.FromLTRB(A.Left, A.Top, A.Left + num, A.Top + num), 180f, 90f);
			}
			graphicsPath2.CloseFigure();
			graphicsPath2 = null;
			return graphicsPath;
		}
	}
}
