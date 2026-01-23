using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;

namespace ExcelAddIn1.Library2.Insert;

public sealed class Shapes
{
	[CompilerGenerated]
	internal sealed class PE
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public Worksheet A;

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}

		[SpecialName]
		internal void B()
		{
			this.A.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
	}

	internal static Microsoft.Office.Interop.Excel.Shape A(Microsoft.Office.Interop.PowerPoint.Shape A, Worksheet B)
	{
		clsClipboard.CopyWithWait((Action)([SpecialName] () =>
		{
			A.Copy();
		}), 4000);
		clsClipboard.PasteWithRetries((Action)([SpecialName] () =>
		{
			B.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}), 5, 4000);
		return B.Shapes.Item(B.Shapes.Count);
	}
}
