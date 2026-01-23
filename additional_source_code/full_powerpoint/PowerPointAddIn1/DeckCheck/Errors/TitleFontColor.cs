using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TitleFontColor : BaseColorError
{
	public TitleFontColor(Slide sld, Shape shp, int intColor, ChartTitle title, Severity sev)
		: base(ErrorType.ColorPaletteChartTitleFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartTitle = title;
		((BaseError)this).Title = AH.A(29763);
		((BaseError)this).Subtitle = AH.A(29808);
		((BaseError)this).ProgrammaticFixFailIsLikely = !A(shp);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		PowerPointAddIn1.DeckCheck.Fix.Colors.RecolorChartFont([SpecialName] () => base.ChartTitle.Format, ColorTranslator.ToOle(color));
	}

	private bool A(Shape A)
	{
		return clsCharts.A(A, new int[7] { 123, 120, 118, 122, 121, 117, 140 });
	}

	[SpecialName]
	[CompilerGenerated]
	private ChartFormat A()
	{
		return base.ChartTitle.Format;
	}
}
