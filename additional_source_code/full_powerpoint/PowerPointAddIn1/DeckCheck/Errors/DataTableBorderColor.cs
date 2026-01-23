using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DataTableBorderColor : BaseColorError
{
	public DataTableBorderColor(Slide sld, Shape shp, int intColor, DataTable tbl, Severity sev)
		: base(ErrorType.ColorPaletteChartDataTableBorder, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.DataTable = tbl;
		((BaseError)this).Title = AH.A(25354);
		((BaseError)this).Subtitle = AH.A(25401);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.DataTable.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
