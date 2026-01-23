using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LegendFillColor : BaseColorError
{
	public LegendFillColor(Slide sld, Shape shp, int intColor, Legend legend, Severity sev)
		: base(ErrorType.ColorPaletteChartLegendFill, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = AH.A(27080);
		((BaseError)this).Subtitle = AH.A(27127);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Legend.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
