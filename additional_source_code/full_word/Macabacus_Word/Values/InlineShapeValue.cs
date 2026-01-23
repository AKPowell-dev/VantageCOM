using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Values;

public sealed class InlineShapeValue : IShape
{
	[CompilerGenerated]
	private float A;

	[CompilerGenerated]
	private float B;

	[CompilerGenerated]
	private Range A;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private InlineShape A;

	public float Width
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public float Height
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
	}

	public Range Range
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public int RngStart
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public int RngEnd
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
	}

	public InlineShape InlineShape
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
	}

	public InlineShapeValue(InlineShape shp)
	{
		InlineShape inlineShape = shp;
		A = shp;
		this.A = inlineShape.Width;
		this.B = inlineShape.Height;
		this.A = shp.Range;
		this.A = Range.Start;
		B = Range.End;
		inlineShape = null;
	}

	public int RangeStart()
	{
		return RngStart;
	}

	int IShape.RangeStart()
	{
		//ILSpy generated this explicit interface implementation from .override directive in RangeStart
		return this.RangeStart();
	}

	public int RangeEnd()
	{
		return RngEnd;
	}

	int IShape.RangeEnd()
	{
		//ILSpy generated this explicit interface implementation from .override directive in RangeEnd
		return this.RangeEnd();
	}
}
