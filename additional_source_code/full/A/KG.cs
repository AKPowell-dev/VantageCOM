using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class KG
{
	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private bool m_B;

	[CompilerGenerated]
	private string m_B;

	[CompilerGenerated]
	private double m_A;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private XlUnderlineStyle m_A;

	private bool Bold
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

	private int Color
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

	private string FontStyle
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

	private bool Italic
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private string Name
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private double Size
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

	private bool Subscript
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

	private bool Superscript
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

	private XlUnderlineStyle Underline
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

	internal KG()
	{
	}

	internal void A(ref Font A)
	{
		Font font = A;
		Color = Conversions.ToInteger(font.Color);
		FontStyle = Conversions.ToString(font.FontStyle);
		Bold = Conversions.ToBoolean(font.Bold);
		Underline = (XlUnderlineStyle)Conversions.ToInteger(font.Underline);
		Italic = Conversions.ToBoolean(font.Italic);
		Name = Conversions.ToString(font.Name);
		Size = Conversions.ToDouble(font.Size);
		Subscript = Conversions.ToBoolean(font.Subscript);
		Superscript = Conversions.ToBoolean(font.Superscript);
		font = null;
	}

	internal void B(ref Font A)
	{
		Font obj = A;
		obj.Color = Color;
		obj.FontStyle = FontStyle;
		obj.Bold = Bold;
		obj.Underline = Underline;
		obj.Italic = Italic;
		obj.Name = Name;
		obj.Size = Size;
		obj.Subscript = Subscript;
		obj.Superscript = Superscript;
		_ = null;
	}
}
