using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DataTableFillColor : BaseColorError
{
	public DataTableFillColor(Slide sld, Shape shp, int intColor, DataTable tbl, Severity sev)
		: base(ErrorType.ColorPaletteChartDataTableFill, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.DataTable = tbl;
		((BaseError)this).Title = AH.A(25537);
		((BaseError)this).Subtitle = AH.A(25580);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.DataTable.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
