using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes.Templated;

public sealed class ImageInput : BaseInput
{
	[CompilerGenerated]
	private Shape A;

	[CompilerGenerated]
	private float A;

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

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private XlHAlign A;

	[CompilerGenerated]
	private XlVAlign A;

	public Shape Placeholder
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public float Top
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public float Bottom
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

	public float Left
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

	public float MaxHeight
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

	public float MaxWidth
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

	public int ZOrder
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public XlHAlign HorizontalAlign
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public XlVAlign VerticalAlign
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public ImageInput(string strLabel, DataTemplate template, Shape shp)
		: base(strLabel, template)
	{
		Placeholder = shp;
		Shape shape = shp;
		Top = shape.Top;
		Bottom = shape.Top + shape.Height;
		Left = shape.Left;
		Right = shape.Left + shape.Width;
		MaxHeight = shape.Height;
		MaxWidth = shape.Width;
		ZOrder = shape.ZOrderPosition;
		shape = null;
		Match match = Regex.Match(shp.Name, AH.A(70392), RegexOptions.IgnoreCase);
		string value = match.Groups[1].Value;
		if (Operators.CompareString(value, AH.A(8598), TextCompare: false) != 0)
		{
			if (Operators.CompareString(value, AH.A(70449), TextCompare: false) != 0)
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
				HorizontalAlign = XlHAlign.xlHAlignCenter;
			}
			else
			{
				HorizontalAlign = XlHAlign.xlHAlignRight;
			}
		}
		else
		{
			HorizontalAlign = XlHAlign.xlHAlignLeft;
		}
		string value2 = match.Groups[2].Value;
		if (Operators.CompareString(value2, AH.A(70454), TextCompare: false) != 0)
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
			if (Operators.CompareString(value2, AH.A(70459), TextCompare: false) != 0)
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
				VerticalAlign = XlVAlign.xlVAlignCenter;
			}
			else
			{
				VerticalAlign = XlVAlign.xlVAlignBottom;
			}
		}
		else
		{
			VerticalAlign = XlVAlign.xlVAlignTop;
		}
		match = null;
	}
}
