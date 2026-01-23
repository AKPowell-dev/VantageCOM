using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class MisalignedShapes
{
	[CompilerGenerated]
	private int A;

	public int Tolerance
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

	public MisalignedShapes()
	{
		Tolerance = Align.GetTolerance();
	}

	public void Check(Slide sld)
	{
		Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = sld.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
		if (shapeRange == null)
		{
			return;
		}
		if (shapeRange.Count >= 3)
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
			list = Align.GetComparisonShapes(shapeRange);
			using List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator = list.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape current = enumerator.Current;
				MsoShapeType type = current.Type;
				if (type == MsoShapeType.msoPlaceholder)
				{
					continue;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				bool flag = false;
				bool flag2 = false;
				bool flag3 = false;
				bool flag4 = false;
				double pos = Math.Round(current.Top, Align.MISALIGNMENT_ROUND);
				flag = Align.IsTopMisaligned(list, ref pos, Tolerance);
				pos = Math.Round(current.Left, Align.MISALIGNMENT_ROUND);
				flag2 = Align.IsLeftMisaligned(list, ref pos, Tolerance);
				if (!flag)
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
					pos = Math.Round(current.Top + current.Height, Align.MISALIGNMENT_ROUND);
					flag3 = Align.IsBottomMisaligned(list, ref pos, Tolerance);
				}
				if (!flag2)
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
					pos = Math.Round(current.Left + current.Width, Align.MISALIGNMENT_ROUND);
					flag4 = Align.IsRightMisaligned(list, ref pos, Tolerance);
				}
				if (!flag)
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
					if (!flag2 && !flag3)
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
						if (!flag4)
						{
							continue;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
					}
				}
				Main.Analysis.Errors.Add(new MisalignedShape(sld, current, Tolerance, flag, flag2, flag3, flag4));
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_01c6;
				}
				continue;
				end_IL_01c6:
				break;
			}
		}
		shapeRange = null;
		list = null;
	}
}
