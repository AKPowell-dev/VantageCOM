using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class MisalignedShape : BaseError
{
	[CompilerGenerated]
	private new bool A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private new int A;

	private bool AlignTop
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

	private bool AlignLeft
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

	private bool AlignBottom
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

	private bool AlignRight
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

	internal int Tolerance
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

	public MisalignedShape(Slide sld, Shape shp, int intTolerance, bool blnTop, bool blnLeft, bool blnBottom, bool blnRight)
		: base(ErrorType.MisalignedShape, Main.Analysis.Options.MisalignedShapes, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(34309);
		((BaseError)this).Subtitle = AH.A(33556) + shp.Name + AH.A(34342);
		Tolerance = intTolerance;
		AlignTop = blnTop;
		AlignLeft = blnLeft;
		AlignBottom = blnBottom;
		AlignRight = blnRight;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		ShapeRange shapes = base.Slide.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
		List<Shape> list = new List<Shape>();
		list = Align.GetComparisonShapes(shapes);
		Shape shape = base.Shape;
		if (AlignTop)
		{
			double pos = Math.Round(shape.Top, Align.MISALIGNMENT_ROUND);
			if (Align.IsTopMisaligned(list, ref pos, Tolerance))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				shape.Top = (float)pos;
			}
		}
		else if (AlignBottom)
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
			double pos = Math.Round(shape.Top + shape.Height, Align.MISALIGNMENT_ROUND);
			if (Align.IsBottomMisaligned(list, ref pos, Tolerance))
			{
				shape.Top = (float)(pos - (double)shape.Height);
			}
		}
		if (AlignLeft)
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
			double pos = Math.Round(shape.Left, Align.MISALIGNMENT_ROUND);
			if (Align.IsLeftMisaligned(list, ref pos, Tolerance))
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
				shape.Left = (float)pos;
			}
		}
		else if (AlignRight)
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
			double pos = Math.Round(shape.Left + shape.Width, Align.MISALIGNMENT_ROUND);
			if (Align.IsRightMisaligned(list, ref pos, Tolerance))
			{
				shape.Left = (float)(pos - (double)shape.Width);
			}
		}
		shape = null;
		list = null;
	}
}
