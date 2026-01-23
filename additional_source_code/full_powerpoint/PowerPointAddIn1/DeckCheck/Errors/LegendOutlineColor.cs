using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LegendOutlineColor : BaseColorError
{
	public LegendOutlineColor(Slide sld, Shape shp, int intColor, Legend legend, Severity sev)
		: base(ErrorType.ColorPaletteChartLegendFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = AH.A(27446);
		((BaseError)this).Subtitle = AH.A(27509);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Legend.Format.TextFrame2.TextRange.Font.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
