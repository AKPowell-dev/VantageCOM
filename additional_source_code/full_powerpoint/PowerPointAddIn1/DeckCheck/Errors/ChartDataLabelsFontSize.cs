using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartDataLabelsFontSize : ChartFontSizeBase
{
	public ChartDataLabelsFontSize(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, float size, int limit, IMsoSeries ser)
		: base(sld, shp, size, limit)
	{
		((BaseError)this).Series = ser;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		try
		{
			((IMsoDataLabels)((BaseError)this).Series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Size = A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
