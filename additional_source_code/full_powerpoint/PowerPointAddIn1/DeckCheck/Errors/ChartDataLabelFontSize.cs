using System;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartDataLabelFontSize : ChartFontSizeBase
{
	private new IMsoDataLabel A;

	public ChartDataLabelFontSize(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, float size, int limit, IMsoDataLabel label)
		: base(sld, shp, size, limit)
	{
		A = label;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		try
		{
			A.Font.Size = base.A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
