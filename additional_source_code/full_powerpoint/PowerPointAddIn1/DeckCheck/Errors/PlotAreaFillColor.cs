using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlotAreaFillColor : BaseColorError
{
	public PlotAreaFillColor(Slide sld, Shape shp, int intColor, PlotArea plot, Severity sev)
		: base(ErrorType.ColorPaletteChartPlotFill, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.PlotArea = plot;
		((BaseError)this).Title = AH.A(28063);
		((BaseError)this).Subtitle = AH.A(28104);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.PlotArea.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
