using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DataTableFontColor : BaseColorError
{
	public DataTableFontColor(Slide sld, Shape shp, int intColor, DataTable tbl, Severity sev)
		: base(ErrorType.ColorPaletteChartDataTableFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.DataTable = tbl;
		((BaseError)this).Title = AH.A(25712);
		((BaseError)this).Subtitle = AH.A(25755);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		PowerPointAddIn1.DeckCheck.Fix.Colors.RecolorChartFont([SpecialName] () => base.DataTable.Format, ColorTranslator.ToOle(color));
	}

	[SpecialName]
	[CompilerGenerated]
	private ChartFormat A()
	{
		return base.DataTable.Format;
	}
}
