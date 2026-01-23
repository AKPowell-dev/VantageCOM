using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartAxisTitleFontSize : ChartFontSizeBase
{
	public ChartAxisTitleFontSize(Slide sld, Shape shp, float size, int limit, AxisTitle axTitle)
		: base(sld, shp, size, limit)
	{
		base.AxisTitle = axTitle;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		try
		{
			base.AxisTitle.Font.Size = A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
