using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingLegendFillColor : BaseColorError
{
	public NonconformingLegendFillColor(object shp, int intColor, Legend legend, Severity sev)
		: base(ErrorType.ColorPaletteChartLegendFill, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = XC.A(33572);
		((BaseError)this).Subtitle = XC.A(33649);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.Legend.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
