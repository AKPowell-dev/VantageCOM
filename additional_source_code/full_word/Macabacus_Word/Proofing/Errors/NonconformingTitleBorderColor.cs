using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingTitleBorderColor : BaseColorError
{
	public NonconformingTitleBorderColor(object shp, int intColor, ChartTitle title, Severity sev)
		: base(ErrorType.ColorPaletteChartTitleBorder, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartTitle = title;
		((BaseError)this).Title = XC.A(34657);
		((BaseError)this).Subtitle = XC.A(34736);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		base.ChartTitle.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
