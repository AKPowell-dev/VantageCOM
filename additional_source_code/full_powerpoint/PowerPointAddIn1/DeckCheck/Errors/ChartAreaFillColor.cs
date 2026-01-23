using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartAreaFillColor : BaseColorError
{
	public ChartAreaFillColor(Slide sld, Shape shp, int intColor, Severity sev)
		: base(ErrorType.ColorPaletteChartAreaFill, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(25179);
		((BaseError)this).Subtitle = AH.A(25222);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Shape.Chart.ChartArea.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
