using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingLegendBorderColor : BaseColorError
{
	public NonconformingLegendBorderColor(object shp, int intColor, Legend legend, Severity sev)
		: base(ErrorType.ColorPaletteChartLegendBorder, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = XC.A(33335);
		((BaseError)this).Subtitle = XC.A(33416);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.Legend.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
