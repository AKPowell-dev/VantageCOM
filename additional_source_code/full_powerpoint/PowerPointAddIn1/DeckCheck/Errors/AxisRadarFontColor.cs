using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AxisRadarFontColor : BaseColorError
{
	public AxisRadarFontColor(Slide sld, Shape shp, int intColor, ChartGroup chtGroup, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartGroup = chtGroup;
		((BaseError)this).Title = AH.A(20539);
		((BaseError)this).Subtitle = AH.A(20582);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.ChartGroup.RadarAxisLabels.Font.Color = ColorTranslator.ToOle(color);
	}
}
