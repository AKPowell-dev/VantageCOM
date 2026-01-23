using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HiLoLinesColor : BaseColorError
{
	public HiLoLinesColor(Slide sld, Shape shp, int intColor, ChartGroup grp, Severity sev)
		: base(ErrorType.ColorPaletteChartHiLoLines, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartGroup = grp;
		((BaseError)this).Title = AH.A(26539);
		((BaseError)this).Subtitle = AH.A(26574);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.ChartGroup.HiLoLines.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
