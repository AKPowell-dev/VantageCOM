using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlotAreaBorderColor : BaseColorError
{
	public PlotAreaBorderColor(Slide sld, Shape shp, int intColor, PlotArea plot, Severity sev)
		: base(ErrorType.ColorPaletteChartPlotBorder, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.PlotArea = plot;
		((BaseError)this).Title = AH.A(27884);
		((BaseError)this).Subtitle = AH.A(27929);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.PlotArea.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
