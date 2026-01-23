using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Shapes;

public sealed class ShapeCellDistance
{
	[CompilerGenerated]
	private Shape A;

	[CompilerGenerated]
	private Cell A;

	[CompilerGenerated]
	private float A;

	internal Shape Shape
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

	internal Cell Cell
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

	internal float Distance
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

	public ShapeCellDistance(Shape shp, Cell cell, float d)
	{
		Shape = shp;
		Cell = cell;
		Distance = d;
	}
}
