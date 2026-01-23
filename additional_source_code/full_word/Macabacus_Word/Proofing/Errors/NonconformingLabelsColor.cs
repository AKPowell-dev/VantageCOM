using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingLabelsColor : BaseColorError
{
	public NonconformingLabelsColor(object shp, int intColor, IMsoSeries series, List<IMsoDataLabel> labels, Severity sev)
		: base(ErrorType.ColorPaletteChartDataLabels, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Series = series;
		((BaseError)this).DataLabels = labels;
		((BaseError)this).Title = XC.A(33114);
		((BaseError)this).Subtitle = XC.A(33187);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		((IMsoDataLabels)((BaseError)this).Series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Color = ColorTranslator.ToOle(color);
		undoRecord.EndCustomRecord();
	}
}
