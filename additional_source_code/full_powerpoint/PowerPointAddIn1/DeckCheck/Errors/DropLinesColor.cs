using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DropLinesColor : BaseColorError
{
	public DropLinesColor(Slide sld, Shape shp, int intColor, ChartGroup grp, Severity sev)
		: base(ErrorType.ColorPaletteChartDropLines, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartGroup = grp;
		((BaseError)this).Title = AH.A(25887);
		((BaseError)this).Subtitle = AH.A(25920);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.ChartGroup.DropLines.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
