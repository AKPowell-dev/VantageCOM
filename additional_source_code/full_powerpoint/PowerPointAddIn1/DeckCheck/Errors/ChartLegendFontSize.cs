using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartLegendFontSize : ChartFontSizeBase
{
	public ChartLegendFontSize(Slide sld, Shape shp, float size, int limit, Legend leg)
		: base(sld, shp, size, limit)
	{
		base.Legend = leg;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		try
		{
			base.Legend.Font.Size = A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
