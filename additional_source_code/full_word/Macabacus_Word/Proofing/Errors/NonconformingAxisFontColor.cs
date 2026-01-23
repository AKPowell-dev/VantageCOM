using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingAxisFontColor : BaseColorError
{
	public NonconformingAxisFontColor(object shp, int intColor, XlAxisGroup axisGroup, Axis axis, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisFont, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Axis = axis;
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
					((BaseError)this).Title = XC.A(28744);
					((BaseError)this).Subtitle = XC.A(28821);
					return;
				}
			}
		}
		((BaseError)this).Title = XC.A(28975);
		((BaseError)this).Subtitle = XC.A(29056);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.Axis.TickLabels.Font.Color = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
