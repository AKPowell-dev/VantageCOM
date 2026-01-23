using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartDataLabelMissing : BaseError
{
	public ChartDataLabelMissing(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, IMsoSeries series, List<IMsoDataLabel> labels)
		: base(ErrorType.ChartDataLabelMissing, ((Settings)Main.Analysis.Options).CheckChartElements, sld, shp, blnHasFix: true)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Series = series;
		((BaseError)this).DataLabels = labels;
		((BaseError)this).Title = AH.A(19029);
		((BaseError)this).Subtitle = AH.A(19098) + series.Name;
		((BaseError)this).Tooltip = AH.A(19117);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		((IMsoDataLabels)((BaseError)this).Series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).ShowValue = true;
	}
}
