using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartTitleFontSize : ChartFontSizeBase
{
	public ChartTitleFontSize(Slide sld, Shape shp, float size, int limit, ChartTitle title)
		: base(sld, shp, size, limit)
	{
		base.ChartTitle = title;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		try
		{
			base.ChartTitle.Font.Size = A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
