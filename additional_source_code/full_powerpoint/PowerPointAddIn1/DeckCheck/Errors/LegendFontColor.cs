using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LegendFontColor : BaseColorError
{
	public LegendFontColor(Slide sld, Shape shp, int intColor, Legend legend, Severity sev)
		: base(ErrorType.ColorPaletteChartLegendFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = AH.A(27263);
		((BaseError)this).Subtitle = AH.A(27310);
		((BaseError)this).ProgrammaticFixFailIsLikely = !A(shp);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		PowerPointAddIn1.DeckCheck.Fix.Colors.RecolorChartFont([SpecialName] () => base.Legend.Format, ColorTranslator.ToOle(color));
	}

	private bool A(Shape A)
	{
		return clsCharts.A(A, new int[3] { 119, 117, 140 });
	}

	[SpecialName]
	[CompilerGenerated]
	private ChartFormat A()
	{
		return base.Legend.Format;
	}
}
