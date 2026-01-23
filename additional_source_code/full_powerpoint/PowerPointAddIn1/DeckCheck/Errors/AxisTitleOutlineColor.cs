using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AxisTitleOutlineColor : BaseColorError
{
	public AxisTitleOutlineColor(Slide sld, Shape shp, int intColor, XlAxisGroup axisGroup, AxisTitle axisTitle, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisTitleFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.AxisTitle = axisTitle;
		if (axisGroup == XlAxisGroup.xlPrimary)
		{
			((BaseError)this).Title = AH.A(22388);
			((BaseError)this).Subtitle = AH.A(22463);
		}
		else
		{
			((BaseError)this).Title = AH.A(22629);
			((BaseError)this).Subtitle = AH.A(22708);
		}
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.AxisTitle.Format.TextFrame2.TextRange.Font.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
