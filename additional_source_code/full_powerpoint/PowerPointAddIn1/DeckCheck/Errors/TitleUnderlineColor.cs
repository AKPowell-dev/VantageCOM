using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TitleUnderlineColor : BaseColorError
{
	public TitleUnderlineColor(Slide sld, Shape shp, int intColor, ChartTitle title, Severity sev)
		: base(ErrorType.ColorPaletteChartTitleFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartTitle = title;
		((BaseError)this).Title = AH.A(30153);
		((BaseError)this).Subtitle = AH.A(30218);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.ChartTitle.Format.TextFrame2.TextRange.Font.UnderlineColor.RGB = ColorTranslator.ToOle(color);
	}
}
