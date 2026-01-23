using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingAxisTitleFontColor : BaseColorError
{
	public NonconformingAxisTitleFontColor(object shp, int intColor, XlAxisGroup axisGroup, AxisTitle axisTitle, Severity sev)
		: base(ErrorType.ColorPaletteChartAxisTitleFont, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.AxisTitle = axisTitle;
		if (axisGroup == XlAxisGroup.xlPrimary)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					((BaseError)this).Title = XC.A(30285);
					((BaseError)this).Subtitle = XC.A(30374);
					return;
				}
			}
		}
		((BaseError)this).Title = XC.A(30540);
		((BaseError)this).Subtitle = XC.A(30633);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		Macabacus_Word.Proofing.Fix.Colors.RecolorChartFont(base.AxisTitle.Format, ColorTranslator.ToOle(color));
		undoRecord.EndCustomRecord();
	}
}
