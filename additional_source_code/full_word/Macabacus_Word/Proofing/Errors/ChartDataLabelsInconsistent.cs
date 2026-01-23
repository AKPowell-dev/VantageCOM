using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class ChartDataLabelsInconsistent : BaseError
{
	public ChartDataLabelsInconsistent(Microsoft.Office.Interop.Word.Shape shp)
		: base(ErrorType.ChartDataLabelsInconsistent, ((Settings)Main.Analysis.Options).CheckChartElements, shp, blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = XC.A(27331);
		((BaseError)this).Subtitle = XC.A(27392);
	}

	public override void FixAction()
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27278));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)base.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
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
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00f8;
				}
				continue;
				end_IL_00f8:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
