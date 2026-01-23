using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AxisBorderColor : BaseColorError
{
	public AxisBorderColor(Slide sld, Shape shp, int intColor, XlAxisGroup axisGroup, Axis axis, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisBorder, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Axis = axis;
		if (axisGroup == XlAxisGroup.xlPrimary)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					((BaseError)this).Title = AH.A(19767);
					((BaseError)this).Subtitle = AH.A(19818);
					return;
				}
			}
		}
		((BaseError)this).Title = AH.A(19960);
		((BaseError)this).Subtitle = AH.A(20015);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Axis.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
