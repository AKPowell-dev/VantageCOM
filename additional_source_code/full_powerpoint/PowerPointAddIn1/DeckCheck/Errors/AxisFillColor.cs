using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AxisFillColor : BaseColorError
{
	public AxisFillColor(Slide sld, Shape shp, int intColor, XlAxisGroup axisGroup, Axis axis, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisFill, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Axis = axis;
		if (axisGroup == XlAxisGroup.xlPrimary)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					((BaseError)this).Title = AH.A(20161);
					((BaseError)this).Subtitle = AH.A(20208);
					return;
				}
			}
		}
		((BaseError)this).Title = AH.A(20346);
		((BaseError)this).Subtitle = AH.A(20397);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Axis.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
