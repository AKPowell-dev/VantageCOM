using System;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Values;

public sealed class ShapeValue : IShape
{
	[CompilerGenerated]
	private Shape A;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private Range A;

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
	private WdRelativeHorizontalPosition A;

	[CompilerGenerated]
	private WdRelativeVerticalPosition A;

	[CompilerGenerated]
	private float F;

	[CompilerGenerated]
	private float G;

	[CompilerGenerated]
	private float H;

	[CompilerGenerated]
	private float I;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private int C;

	[CompilerGenerated]
	private WdWrapType A;

	[CompilerGenerated]
	private WdWrapSideType A;

	[CompilerGenerated]
	private float J;

	[CompilerGenerated]
	private float K;

	[CompilerGenerated]
	private float L;

	[CompilerGenerated]
	private float M;

	[CompilerGenerated]
	private bool C;

	public Shape Shape
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public bool IsInHeaderFooter
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public bool IsInsideGroup
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
	}

	public Range Anchor
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public float Top
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public float Left
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
	}

	public float Width
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
	}

	public float Height
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
	}

	public float Rotation
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
	}

	public WdRelativeHorizontalPosition RelativeHorizontalPosition
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public WdRelativeVerticalPosition RelativeVerticalPosition
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public float WidthRelative
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
	}

	public float HeightRelative
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
	}

	public float LeftRelative
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
	}

	public float TopRelative
	{
		[CompilerGenerated]
		get
		{
			return I;
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

	public int ZOrderPosition
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
	}

	public WdWrapType WrapType
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public WdWrapSideType WrapSide
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
	}

	public float DistanceTop
	{
		[CompilerGenerated]
		get
		{
			return J;
		}
	}

	public float DistanceBottom
	{
		[CompilerGenerated]
		get
		{
			return K;
		}
	}

	public float DistanceLeft
	{
		[CompilerGenerated]
		get
		{
			return L;
		}
	}

	public float DistanceRight
	{
		[CompilerGenerated]
		get
		{
			return M;
		}
	}

	public bool AllowOverlap
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
	}

	public ShapeValue(Shape shp)
		: this(shp, isInsideGroup: false)
	{
	}

	public ShapeValue(Shape shp, bool isInsideGroup)
	{
		Shape shape = shp;
		this.A = shp;
		this.A = shape.Anchor;
		this.A = clsUtilities.IsRangeInHeaderFooter(Anchor);
		this.B = isInsideGroup;
		this.C = shape.Width;
		D = shape.Height;
		this.A = shape.Top;
		this.B = shape.Left;
		this.A = Anchor.Start;
		B = Anchor.End;
		E = shape.Rotation;
		this.C = shape.ZOrderPosition;
		try
		{
			this.A = shape.RelativeVerticalPosition;
			this.A = shape.RelativeHorizontalPosition;
			I = shape.TopRelative;
			H = shape.LeftRelative;
			G = shape.HeightRelative;
			F = shape.WidthRelative;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		WrapFormat wrapFormat = shape.WrapFormat;
		this.A = wrapFormat.Type;
		A = wrapFormat.Side;
		J = wrapFormat.DistanceTop;
		K = wrapFormat.DistanceBottom;
		L = wrapFormat.DistanceLeft;
		M = wrapFormat.DistanceRight;
		C = wrapFormat.AllowOverlap != 0;
		wrapFormat = null;
		shape = null;
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
