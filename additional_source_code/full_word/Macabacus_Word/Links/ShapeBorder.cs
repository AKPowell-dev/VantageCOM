using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Links;

public sealed class ShapeBorder
{
	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private MsoLineStyle A;

	[CompilerGenerated]
	private MsoTriState A;

	[CompilerGenerated]
	private float A;

	public int Color
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

	public MsoLineStyle Style
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

	public float Weight
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

	public ShapeBorder(Microsoft.Office.Interop.Word.Shape shape)
	{
		Microsoft.Office.Interop.Word.LineFormat line = shape.Line;
		Visible = line.Visible;
		Color = line.ForeColor.RGB;
		Style = line.Style;
		Weight = line.Weight;
		line = null;
	}
}
