using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartAreaBorderColor : BaseColorError
{
	public ChartAreaBorderColor(Slide sld, Shape shp, int intColor, Severity sev)
		: base(ErrorType.ColorPaletteChartAreaBorder, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(24996);
		((BaseError)this).Subtitle = AH.A(25043);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Shape.Chart.ChartArea.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
