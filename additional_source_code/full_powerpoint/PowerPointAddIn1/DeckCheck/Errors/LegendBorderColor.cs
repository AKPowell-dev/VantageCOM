using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LegendBorderColor : BaseColorError
{
	public LegendBorderColor(Slide sld, Shape shp, int intColor, Legend legend, Severity sev)
		: base(ErrorType.ColorPaletteChartLegendBorder, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = AH.A(26889);
		((BaseError)this).Subtitle = AH.A(26940);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Legend.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
