using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingLegendFontColor : BaseColorError
{
	public NonconformingLegendFontColor(object shp, int intColor, Legend legend, Severity sev)
		: base(ErrorType.ColorPaletteChartLegendFont, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = XC.A(33801);
		((BaseError)this).Subtitle = XC.A(33878);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		Macabacus_Word.Proofing.Fix.Colors.RecolorChartFont(base.Legend.Format, ColorTranslator.ToOle(color));
		undoRecord.EndCustomRecord();
	}
}
