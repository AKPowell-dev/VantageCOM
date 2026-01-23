using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartTickLabelsFontSize : ChartFontSizeBase
{
	public ChartTickLabelsFontSize(Slide sld, Shape shp, float size, int limit, Axis ax)
		: base(sld, shp, size, limit)
	{
		base.Axis = ax;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		try
		{
			base.Axis.TickLabels.Font.Size = A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
