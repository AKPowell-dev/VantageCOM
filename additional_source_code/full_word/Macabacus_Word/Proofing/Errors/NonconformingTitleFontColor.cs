using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingTitleFontColor : BaseColorError
{
	public NonconformingTitleFontColor(object shp, int intColor, ChartTitle title, Severity sev)
		: base(ErrorType.ColorPaletteChartTitleFont, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.ChartTitle = title;
		((BaseError)this).Title = XC.A(35115);
		((BaseError)this).Subtitle = XC.A(35190);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		Macabacus_Word.Proofing.Fix.Colors.RecolorChartFont(base.ChartTitle.Format, ColorTranslator.ToOle(color));
		undoRecord.EndCustomRecord();
	}
}
