using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Links;

public sealed class InlineShapeBorder
{
	[CompilerGenerated]
	private WdLineStyle A;

	[CompilerGenerated]
	private WdLineWidth A;

	[CompilerGenerated]
	private WdColor A;

	[CompilerGenerated]
	private MsoTriState A;

	[CompilerGenerated]
	private bool A;

	public WdLineStyle OutsideLineStyle
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

	public WdLineWidth OutsideLineWidth
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

	public WdColor OutsideColor
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

	public MsoTriState Visible
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

	public bool Enable
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

	public InlineShapeBorder(InlineShape shape)
	{
		Microsoft.Office.Interop.Word.LineFormat line = shape.Line;
		Visible = line.Visible;
		line = null;
		Enable = shape.Range.Borders.Enable != 0;
		Borders borders = shape.Borders;
		OutsideLineStyle = borders.OutsideLineStyle;
		OutsideColor = borders.OutsideColor;
		OutsideLineWidth = borders.OutsideLineWidth;
		borders = null;
	}
}
