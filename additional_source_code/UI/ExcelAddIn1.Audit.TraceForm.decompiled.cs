using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;

namespace ExcelAddIn1.Audit;

public class TraceForm : Form
{
	public struct VirtualBorder
	{
		public Rectangle Rectangle1;

		public Rectangle Rectangle2;

		public Rectangle Rectangle3;

		public Rectangle Rectangle4;

		public Rectangle Rectangle5;

		public Rectangle Rectangle6;

		public Rectangle Rectangle7;

		public Rectangle Rectangle8;

		public Rectangle Rectangle9;
	}

	private readonly Color m_A;

	private readonly Color m_B;

	private readonly Color C;

	private readonly Color D;

	[CompilerGenerated]
	private bool m_A;

	private VirtualBorder m_A;

	private bool m_B;

	private bool C;

	private bool D;

	private bool E;

	private bool F;

	private bool G;

	private bool H;

	private bool I;

	private bool J;

	private bool K;

	private Point m_A;

	private Point m_B;

	private bool L;

	private bool IsActivated
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

	[Browsable(false)]
	[ReadOnly(true)]
	public override Color BackColor
	{
		get
		{
			return base.BackColor;
		}
		set
		{
			base.BackColor = value;
		}
	}

	[Browsable(false)]
	[ReadOnly(true)]
	public override Color ForeColor
	{
		get
		{
			return base.ForeColor;
		}
		set
		{
			base.ForeColor = value;
		}
	}

	[ReadOnly(false)]
	[Browsable(true)]
	public new Font Font
	{
		get
		{
			return base.Font;
		}
		set
		{
			base.Font = value;
		}
	}

	public TraceForm()
	{
		this.m_A = Color.FromArgb(0, 191, 143);
		this.m_B = Color.FromArgb(0, 165, 122);
		this.C = Color.FromArgb(200, 200, 200);
		this.D = Color.FromArgb(175, 175, 175);
		L = true;
		base.AutoScaleMode = AutoScaleMode.Font;
		DoubleBuffered = true;
		Font = new Font(VH.A(50021), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		base.FormBorderStyle = FormBorderStyle.None;
		base.MaximizeBox = false;
		base.MinimizeBox = false;
		base.ShowIcon = false;
		base.ShowInTaskbar = false;
		base.TopMost = true;
		base.StartPosition = FormStartPosition.Manual;
		BackColor = this.m_B;
	}

	protected override void OnPaint(PaintEventArgs e)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Expected O, but got Unknown
		base.OnPaint(e);
		clsDisplay val = new clsDisplay();
		Color b;
		if (!IsActivated)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!base.DesignMode)
			{
				b = this.m_B;
				goto IL_004e;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		b = this.m_B;
		goto IL_004e;
		IL_004e:
		checked
		{
			Rectangle rect = new Rectangle(1, 1, base.Width - 2, base.Height - 2);
			Pen pen = new Pen(b, (float)(2.0 * val.Y));
			try
			{
				e.Graphics.DrawRectangle(pen, rect);
			}
			finally
			{
				if (pen != null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						((IDisposable)pen).Dispose();
						break;
					}
				}
			}
			rect = default(Rectangle);
			Font font = new Font(VH.A(50021), 11f);
			StringFormat stringFormat = new StringFormat();
			stringFormat.Alignment = StringAlignment.Near;
			SolidBrush solidBrush = new SolidBrush(Color.White);
			try
			{
				e.Graphics.DrawString(Text, font, solidBrush, new Rectangle(11, 6, base.Width, font.Height + 12), stringFormat);
			}
			finally
			{
				if (solidBrush != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						((IDisposable)solidBrush).Dispose();
						break;
					}
				}
			}
			font.Dispose();
			A();
		}
	}

	protected override void OnActivated(EventArgs e)
	{
		base.OnActivated(e);
		IsActivated = true;
	}

	protected override void OnDeactivate(EventArgs e)
	{
		base.OnDeactivate(e);
		IsActivated = false;
	}

	private void A()
	{
		Rectangle clientRectangle = base.ClientRectangle;
		this.m_A = default(VirtualBorder);
		ref VirtualBorder a = ref this.m_A;
		a.Rectangle1 = new Rectangle(new Point(clientRectangle.X, clientRectangle.Y), new Size(4, 4));
		checked
		{
			a.Rectangle2 = new Rectangle(new Point(clientRectangle.X + a.Rectangle1.Width, clientRectangle.Y), new Size(clientRectangle.Width - a.Rectangle1.Width * 2, a.Rectangle1.Height));
			a.Rectangle3 = new Rectangle(new Point(clientRectangle.X + a.Rectangle1.Width + a.Rectangle2.Width, clientRectangle.Y), new Size(4, 4));
			a.Rectangle4 = new Rectangle(new Point(clientRectangle.X, clientRectangle.Y + a.Rectangle1.Height), new Size(a.Rectangle1.Width, clientRectangle.Height - a.Rectangle1.Width * 2));
			a.Rectangle5 = new Rectangle(new Point(clientRectangle.X + a.Rectangle4.Width, clientRectangle.Y + a.Rectangle1.Height), new Size(a.Rectangle2.Width, a.Rectangle4.Height));
			a.Rectangle6 = new Rectangle(new Point(clientRectangle.X + a.Rectangle4.Width + a.Rectangle5.Width, clientRectangle.Y + a.Rectangle1.Height), new Size(a.Rectangle3.Width, a.Rectangle4.Height));
			a.Rectangle7 = new Rectangle(new Point(clientRectangle.X, clientRectangle.Y + a.Rectangle1.Height + a.Rectangle4.Height), new Size(4, 4));
			a.Rectangle8 = new Rectangle(new Point(clientRectangle.X + a.Rectangle7.Width, clientRectangle.Y + a.Rectangle1.Height + a.Rectangle4.Height), new Size(clientRectangle.Width - a.Rectangle7.Width * 2, a.Rectangle7.Height));
			a.Rectangle9 = new Rectangle(new Point(clientRectangle.X + a.Rectangle7.Width + a.Rectangle8.Width, clientRectangle.Y + a.Rectangle1.Height + a.Rectangle4.Height), new Size(4, 4));
			clientRectangle = default(Rectangle);
		}
	}

	private Cursor A(Point A)
	{
		ref VirtualBorder a = ref this.m_A;
		if (a.Rectangle1.Contains(A))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return Cursors.SizeNWSE;
				}
			}
		}
		if (a.Rectangle2.Contains(A))
		{
			return Cursors.SizeNS;
		}
		if (a.Rectangle3.Contains(A))
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return Cursors.SizeNESW;
				}
			}
		}
		if (a.Rectangle4.Contains(A))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return Cursors.SizeWE;
				}
			}
		}
		if (a.Rectangle5.Contains(A))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return Cursors.Default;
				}
			}
		}
		if (a.Rectangle6.Contains(A))
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return Cursors.SizeWE;
				}
			}
		}
		if (a.Rectangle7.Contains(A))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return Cursors.SizeNESW;
				}
			}
		}
		if (a.Rectangle8.Contains(A))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return Cursors.SizeNS;
				}
			}
		}
		if (a.Rectangle9.Contains(A))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return Cursors.SizeNWSE;
				}
			}
		}
		return Cursors.Default;
	}

	protected override void OnMouseMove(MouseEventArgs e)
	{
		base.OnMouseMove(e);
		Point a = PointToScreen(new Point(e.X, e.Y));
		checked
		{
			if (base.Capture)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (this.m_B)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									C = false;
									base.Location = new Point(Control.MousePosition.X - this.m_B.X, Control.MousePosition.Y - this.m_B.Y);
									return;
								}
							}
						}
						if (C)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
								{
									this.m_B = false;
									Rectangle bounds = base.Bounds;
									if (D)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											break;
										}
										base.Bounds = new Rectangle(bounds.X + a.X - this.m_A.X, bounds.Y + a.Y - this.m_A.Y, bounds.Width - a.X + this.m_A.X, bounds.Height - a.Y + this.m_A.Y);
									}
									else if (E)
									{
										base.Bounds = new Rectangle(bounds.X, bounds.Y + a.Y - this.m_A.Y, bounds.Width, bounds.Height - a.Y + this.m_A.Y);
									}
									else if (F)
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
										base.Bounds = new Rectangle(bounds.X, bounds.Y + a.Y - this.m_A.Y, bounds.Width + a.X - this.m_A.X, bounds.Height - a.Y + this.m_A.Y);
									}
									else if (G)
									{
										base.Bounds = new Rectangle(bounds.X + a.X - this.m_A.X, bounds.Y, bounds.Width - a.X + this.m_A.X, bounds.Height);
									}
									else if (H)
									{
										base.Bounds = new Rectangle(bounds.X, bounds.Y, bounds.Width + a.X - this.m_A.X, bounds.Height);
									}
									else if (I)
									{
										base.Bounds = new Rectangle(bounds.X + a.X - this.m_A.X, bounds.Y, bounds.Width - a.X + this.m_A.X, bounds.Height + a.Y - this.m_A.Y);
									}
									else if (J)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											break;
										}
										base.Bounds = new Rectangle(bounds.X, bounds.Y, bounds.Width, bounds.Height + a.Y - this.m_A.Y);
									}
									else if (K)
									{
										base.Bounds = new Rectangle(bounds.X, bounds.Y, bounds.Width + a.X - this.m_A.X, bounds.Height + a.Y - this.m_A.Y);
									}
									this.m_A = a;
									Invalidate();
									return;
								}
								}
							}
						}
						return;
					}
				}
			}
			this.m_B = new Point(e.X, e.Y);
			Cursor = A(this.m_B);
		}
	}

	protected override void OnMouseDown(MouseEventArgs e)
	{
		base.OnMouseDown(e);
		if (e.Button != MouseButtons.Left)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Point point = new Point(e.X, e.Y);
			if (this.m_A.Rectangle1.Contains(this.m_B))
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
				C = true;
				D = true;
				E = false;
				F = false;
				G = false;
				H = false;
				I = false;
				J = false;
				K = false;
				this.m_A = PointToScreen(point);
			}
			else if (this.m_A.Rectangle2.Contains(this.m_B))
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
				C = true;
				D = false;
				E = true;
				F = false;
				G = false;
				H = false;
				I = false;
				J = false;
				K = false;
				this.m_A = PointToScreen(point);
			}
			else if (this.m_A.Rectangle3.Contains(this.m_B))
			{
				C = true;
				D = false;
				E = false;
				F = true;
				G = false;
				H = false;
				I = false;
				J = false;
				K = false;
				this.m_A = PointToScreen(point);
			}
			else if (this.m_A.Rectangle4.Contains(this.m_B))
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
				C = true;
				D = false;
				E = false;
				F = false;
				G = true;
				H = false;
				I = false;
				J = false;
				K = false;
				this.m_A = PointToScreen(point);
			}
			else if (this.m_A.Rectangle5.Contains(this.m_B))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				this.m_B = true;
				C = false;
				D = false;
				E = false;
				F = false;
				G = false;
				H = false;
				I = false;
				J = false;
				K = false;
				this.m_B = point;
				Cursor = Cursors.SizeAll;
			}
			else if (this.m_A.Rectangle6.Contains(this.m_B))
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
				C = true;
				D = false;
				E = false;
				F = false;
				G = false;
				H = true;
				I = false;
				J = false;
				K = false;
				this.m_A = PointToScreen(point);
			}
			else if (this.m_A.Rectangle7.Contains(this.m_B))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				C = true;
				D = false;
				E = false;
				F = false;
				G = false;
				H = false;
				I = true;
				J = false;
				K = false;
				this.m_A = PointToScreen(point);
			}
			else if (this.m_A.Rectangle8.Contains(this.m_B))
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
				C = true;
				D = false;
				E = false;
				F = false;
				G = false;
				H = false;
				I = false;
				J = true;
				K = false;
				this.m_A = PointToScreen(point);
			}
			else if (this.m_A.Rectangle9.Contains(this.m_B))
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
				C = true;
				D = false;
				E = false;
				F = false;
				G = false;
				H = false;
				I = false;
				J = false;
				K = true;
				this.m_A = PointToScreen(point);
			}
			point = default(Point);
			return;
		}
	}

	protected override void OnMouseUp(MouseEventArgs e)
	{
		base.OnMouseUp(e);
		this.m_B = false;
		C = false;
	}

	private void B()
	{
		SuspendLayout();
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		base.ClientSize = new Size(284, 261);
		base.Name = VH.A(50038);
		ResumeLayout(performLayout: false);
	}

	public void SetPadding()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Expected O, but got Unknown
		clsDisplay val = new clsDisplay();
		base.Padding = checked(new Padding((int)Math.Round(2.0 * val.X), (int)Math.Round(30.0 * val.Y), (int)Math.Round(2.0 * val.X), (int)Math.Round(2.0 * val.Y)));
		val = null;
	}
}
