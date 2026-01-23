using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AxisTitleUnderlineColor : BaseColorError
{
	public AxisTitleUnderlineColor(Slide sld, Shape shp, int intColor, XlAxisGroup axisGroup, AxisTitle axisTitle, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisTitleFont, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.AxisTitle = axisTitle;
		if (axisGroup == XlAxisGroup.xlPrimary)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					((BaseError)this).Title = AH.A(22878);
					((BaseError)this).Subtitle = AH.A(22957);
					return;
				}
			}
		}
		((BaseError)this).Title = AH.A(23127);
		((BaseError)this).Subtitle = AH.A(23210);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.AxisTitle.Format.TextFrame2.TextRange.Font.UnderlineColor.RGB = ColorTranslator.ToOle(color);
	}
}
