using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartDataTableFontSize : ChartFontSizeBase
{
	public ChartDataTableFontSize(Slide sld, Shape shp, float size, int limit, DataTable tbl)
		: base(sld, shp, size, limit)
	{
		base.DataTable = tbl;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		try
		{
			base.DataTable.Font.Size = A;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
