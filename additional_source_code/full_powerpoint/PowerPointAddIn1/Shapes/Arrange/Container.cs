using System;
using System.Drawing;
using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.Shapes.Arrange;

public sealed class Container
{
	[CompilerGenerated]
	private float m_A;

	[CompilerGenerated]
	private float B;

	[CompilerGenerated]
	private float C;

	[CompilerGenerated]
	private float D;

	[CompilerGenerated]
	private float E;

	[CompilerGenerated]
	private float F;

	public float Top
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

	public float Left
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public float Bottom
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public float Right
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	public float Height
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	public float Width
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	public Container(Preferences prefs, RectangleF rect)
	{
		Width = rect.Width;
		Height = rect.Height;
		Left = rect.Left;
		Top = rect.Top;
		A(prefs.ContainerPadding);
	}

	private void A(float A)
	{
		if (A > 0f)
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
			Width = Math.Max(0f, Width - 2f * A);
			Height = Math.Max(0f, Height - 2f * A);
			Left += A;
			Top += A;
		}
		Right = Left + Width;
		Bottom = Top + Height;
	}
}
