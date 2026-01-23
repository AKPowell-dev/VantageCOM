using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;

namespace A;

internal sealed class HD
{
	[CompilerGenerated]
	internal sealed class GD
	{
		public Shape A;

		public ShapeRange A;

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}

		[SpecialName]
		internal void B()
		{
			this.A.Copy();
		}
	}

	internal static void A(Presentation A)
	{
		Shape A2;
		ShapeRange A3;
		try
		{
			Slide slide = A.Slides[1];
			Shapes shapes = slide.Shapes;
			if (shapes.Count < 2)
			{
				A2 = slide.Shapes[1];
				clsClipboard.CopyWithWait((Action)([SpecialName] () =>
				{
					A2.Copy();
				}), 4000);
			}
			else
			{
				A3 = shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
				clsClipboard.CopyWithWait((Action)([SpecialName] () =>
				{
					A3.Copy();
				}), 4000);
			}
		}
		finally
		{
			A3 = null;
			A2 = null;
			Shapes shapes = null;
			Slide slide = null;
		}
	}
}
