using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Links;

public sealed class TableBorder
{
	[CompilerGenerated]
	private WdColor A;

	[CompilerGenerated]
	private WdColorIndex A;

	[CompilerGenerated]
	private WdLineStyle A;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private WdLineWidth A;

	public WdColor Color
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

	public WdColorIndex ColorIndex
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

	public WdLineStyle Style
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

	public bool Visible
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

	public WdLineWidth Weight
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

	public TableBorder(Border border)
	{
		Border border2 = border;
		Color = border2.Color;
		ColorIndex = border2.ColorIndex;
		Style = border2.LineStyle;
		Weight = border2.LineWidth;
		Visible = border2.Visible;
		border2 = null;
	}
}
