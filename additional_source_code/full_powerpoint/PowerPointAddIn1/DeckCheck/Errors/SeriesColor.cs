using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SeriesColor : BaseColorError
{
	public SeriesColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, IMsoSeries series, Severity sev)
		: base(ErrorType.ColorPaletteChartSeries, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Series = series;
		((BaseError)this).Title = AH.A(28571);
		((BaseError)this).Subtitle = AH.A(28596);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		Charts.A(((BaseError)this).Series, ColorTranslator.ToOle(color));
	}
}
