using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ChartDataLabelsInconsistent : BaseError
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

	public ChartDataLabelsInconsistent(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
		: base(ErrorType.ChartDataLabelsInconsistent, ((Settings)Main.Analysis.Options).CheckChartElements, sld, shp, blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(19421);
		((BaseError)this).Subtitle = AH.A(19482);
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
		IEnumerator enumerator = ((IEnumerable)base.Shape.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
				if (msoSeries.HasDataLabels)
				{
					continue;
				}
				while (true)
				{
					switch (3)
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
				msoSeries.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_012b;
				}
				continue;
				end_IL_012b:
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
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
