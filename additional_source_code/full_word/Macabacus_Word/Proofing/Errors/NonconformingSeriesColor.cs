using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingSeriesColor : BaseColorError
{
	public NonconformingSeriesColor(object shp, int intColor, IMsoSeries series, Severity sev)
		: base(ErrorType.ColorPaletteChartSeries, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Series = series;
		((BaseError)this).Title = XC.A(34472);
		((BaseError)this).Subtitle = XC.A(34527);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		((BaseError)this).Series.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
