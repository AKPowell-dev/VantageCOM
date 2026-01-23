using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingAxisFillColor : BaseColorError
{
	public NonconformingAxisFillColor(object shp, int intColor, XlAxisGroup axisGroup, Axis axis, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisFill, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
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
					((BaseError)this).Title = XC.A(28274);
					((BaseError)this).Subtitle = XC.A(28351);
					return;
				}
			}
		}
		((BaseError)this).Title = XC.A(28505);
		((BaseError)this).Subtitle = XC.A(28586);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.Axis.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
