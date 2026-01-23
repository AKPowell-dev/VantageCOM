using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class ChartDataLabelMissing : BaseError
{
	public ChartDataLabelMissing(Microsoft.Office.Interop.Word.Shape shp, IMsoSeries series, List<IMsoDataLabel> labels)
		: base(ErrorType.ChartDataLabelMissing, ((Settings)Main.Analysis.Options).CheckChartElements, shp, blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Series = series;
		((BaseError)this).DataLabels = labels;
		((BaseError)this).Title = XC.A(26860);
		((BaseError)this).Subtitle = XC.A(26929) + series.Name;
		((BaseError)this).Tooltip = XC.A(26948);
	}

	public override void FixAction()
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(26813));
		((IMsoDataLabels)((BaseError)this).Series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).ShowValue = true;
		undoRecord.EndCustomRecord();
	}
}
