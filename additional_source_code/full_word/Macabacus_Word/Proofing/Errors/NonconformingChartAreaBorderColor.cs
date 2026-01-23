using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingChartAreaBorderColor : BaseColorError
{
	public NonconformingChartAreaBorderColor(object shp, int intColor, Severity sev)
		: base(ErrorType.ColorPaletteChartAreaBorder, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = XC.A(31004);
		((BaseError)this).Subtitle = XC.A(31081);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.Shape.Chart.ChartArea.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
