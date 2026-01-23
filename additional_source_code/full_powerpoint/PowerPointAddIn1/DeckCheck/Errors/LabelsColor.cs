using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LabelsColor : BaseColorError
{
	public LabelsColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, IMsoSeries series, List<IMsoDataLabel> labels, Severity sev)
		: base(ErrorType.ColorPaletteChartDataLabels, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Series = series;
		((BaseError)this).DataLabels = labels;
		((BaseError)this).Title = AH.A(26714);
		((BaseError)this).Subtitle = AH.A(26757);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		((IMsoDataLabels)((BaseError)this).Series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Color = ColorTranslator.ToOle(color);
	}
}
