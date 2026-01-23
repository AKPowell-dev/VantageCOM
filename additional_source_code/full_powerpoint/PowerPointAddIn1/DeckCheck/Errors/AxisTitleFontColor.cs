using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AxisTitleFontColor : BaseColorError
{
	public AxisTitleFontColor(Slide sld, Shape shp, int intColor, XlAxisGroup axisGroup, AxisTitle axisTitle, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisTitleFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.AxisTitle = axisTitle;
		if (axisGroup == XlAxisGroup.xlPrimary)
		{
			((BaseError)this).Title = AH.A(21962);
			((BaseError)this).Subtitle = AH.A(22021);
		}
		else
		{
			((BaseError)this).Title = AH.A(22171);
			((BaseError)this).Subtitle = AH.A(22234);
		}
		((BaseError)this).ProgrammaticFixFailIsLikely = !A(shp);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		PowerPointAddIn1.DeckCheck.Fix.Colors.RecolorChartFont([SpecialName] () => base.AxisTitle.Format, ColorTranslator.ToOle(color));
	}

	private bool A(Shape A)
	{
		return clsCharts.A(A, new int[1] { 119 });
	}

	[SpecialName]
	[CompilerGenerated]
	private ChartFormat A()
	{
		return base.AxisTitle.Format;
	}
}
