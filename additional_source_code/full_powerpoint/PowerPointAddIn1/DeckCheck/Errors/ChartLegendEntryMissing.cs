using System;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartLegendEntryMissing : BaseError
{
	[CompilerGenerated]
	private new int A;

	internal int RequiredUndoSteps
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public ChartLegendEntryMissing(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, string strSubtitle, Legend legend)
		: base(ErrorType.ChartLegendEntryMissing, ((Settings)Main.Analysis.Options).CheckChartElements, sld, shp, blnHasFix: true)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		base.Legend = legend;
		((BaseError)this).Title = AH.A(19591);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(19654);
	}

	public override void FixAction()
	{
		int num = 0;
		CommandBarControl instance = default(CommandBarControl);
		try
		{
			instance = Charts.UndoControl();
			num = Conversions.ToInteger(NewLateBinding.LateGet(instance, null, AH.A(19222), new object[0], null, null, null));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		NG.A.Application.StartNewUndoEntry();
		Chart chart = base.Shape.Chart;
		Legend legend = chart.Legend;
		XlLegendPosition position = legend.Position;
		double top = legend.Top;
		double left = legend.Left;
		_ = null;
		chart.HasLegend = false;
		chart.HasLegend = true;
		Legend legend2 = chart.Legend;
		if (position != XlLegendPosition.xlLegendPositionCustom)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			legend2.Position = position;
		}
		else
		{
			if (left > base.Shape.Chart.ChartArea.Width * 0.5)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				legend2.Position = XlLegendPosition.xlLegendPositionRight;
			}
			else if (top > base.Shape.Chart.ChartArea.Height * 0.5)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				legend2.Position = XlLegendPosition.xlLegendPositionBottom;
			}
			else if (left < base.Shape.Chart.PlotArea.Left)
			{
				legend2.Position = XlLegendPosition.xlLegendPositionLeft;
			}
			else
			{
				legend2.Position = XlLegendPosition.xlLegendPositionTop;
			}
			legend2.Top = top;
			legend2.Left = left;
		}
		legend2 = null;
		_ = null;
		try
		{
			RequiredUndoSteps = Conversions.ToInteger(Operators.SubtractObject(NewLateBinding.LateGet(instance, null, AH.A(19222), new object[0], null, null, null), num));
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		instance = null;
	}
}
