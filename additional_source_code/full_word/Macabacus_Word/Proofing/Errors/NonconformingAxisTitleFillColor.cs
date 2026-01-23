using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingAxisTitleFillColor : BaseColorError
{
	public NonconformingAxisTitleFillColor(object shp, int intColor, XlAxisGroup axisGroup, AxisTitle axisTitle, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisTitleFill, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.AxisTitle = axisTitle;
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
					((BaseError)this).Title = XC.A(29767);
					((BaseError)this).Subtitle = XC.A(29856);
					return;
				}
			}
		}
		((BaseError)this).Title = XC.A(30022);
		((BaseError)this).Subtitle = XC.A(30115);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(29748));
		base.AxisTitle.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
