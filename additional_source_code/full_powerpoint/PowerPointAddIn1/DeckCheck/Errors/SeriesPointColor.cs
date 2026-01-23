using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SeriesPointColor : BaseColorError
{
	public SeriesPointColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, ChartPoint point, Severity sev)
		: base(ErrorType.ColorPaletteChartSeries, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).ChartPoint = point;
		((BaseError)this).Title = AH.A(28409);
		((BaseError)this).Subtitle = AH.A(28446);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		((BaseError)this).ChartPoint.Interior.Color = ColorTranslator.ToOle(color);
	}
}
