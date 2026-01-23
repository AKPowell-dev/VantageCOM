using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartDataLabelNumberFormats : BaseError
{
	[CompilerGenerated]
	private new List<string> A;

	[CompilerGenerated]
	private new int A;

	private List<string> FixOptions
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

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

	public ChartDataLabelNumberFormats(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<string> listLabels, string strSubtitle, List<string> listFixes, PlotArea plot)
		: base(ErrorType.ChartDataLabelNumberFormats, ((Settings)Main.Analysis.Options).CheckChartElements, sld, shp, blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		base.PlotArea = plot;
		((BaseError)this).Title = AH.A(19241);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(19304);
	}

	public override void FixAction(int i)
	{
		string numberFormat = FixOptions[i];
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
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)base.Shape.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
				if (!msoSeries.HasDataLabels)
				{
					continue;
				}
				while (true)
				{
					switch (1)
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
				try
				{
					enumerator2 = ((IEnumerable)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator2.Current;
						if (!msoDataLabel.ShowValue)
						{
							continue;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
						msoDataLabel.NumberFormat = numberFormat;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_014b;
				}
				continue;
				end_IL_014b:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
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
