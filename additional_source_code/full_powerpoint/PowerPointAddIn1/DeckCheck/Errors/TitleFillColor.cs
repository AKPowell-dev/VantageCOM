using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TitleFillColor : BaseColorError
{
	public TitleFillColor(Slide sld, Shape shp, int intColor, ChartTitle title, Severity sev)
		: base(ErrorType.ColorPaletteChartTitleFill, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartTitle = title;
		((BaseError)this).Title = AH.A(29584);
		((BaseError)this).Subtitle = AH.A(29629);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.ChartTitle.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
