using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AxisTitleBorderColor : BaseColorError
{
	public AxisTitleBorderColor(Slide sld, Shape shp, int intColor, XlAxisGroup axisGroup, AxisTitle axisTitle, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisTitleBorder, sev, sld, shp, intColor)
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
					((BaseError)this).Title = AH.A(21094);
					((BaseError)this).Subtitle = AH.A(21157);
					return;
				}
			}
		}
		((BaseError)this).Title = AH.A(21311);
		((BaseError)this).Subtitle = AH.A(21378);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.AxisTitle.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
