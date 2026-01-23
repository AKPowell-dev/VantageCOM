using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class ChartDataLabelNumberFormats : BaseError
{
	private List<string> A;

	private List<string> FixOptions
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public ChartDataLabelNumberFormats(Microsoft.Office.Interop.Word.Shape shp, List<string> listLabels, string strSubtitle, List<string> listFixes)
		: base(ErrorType.ChartDataLabelNumberFormats, ((Settings)Main.Analysis.Options).CheckChartElements, shp, blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = XC.A(27098);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = XC.A(27161);
	}

	public override void FixAction(int i)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27053));
		string numberFormat = FixOptions[i];
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)base.Chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				IMsoSeries msoSeries = (IMsoSeries)enumerator.Current;
				if (!msoSeries.HasDataLabels)
				{
					continue;
				}
				{
					enumerator2 = ((IEnumerable)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator2.Current;
							if (msoDataLabel.ShowValue)
							{
								msoDataLabel.NumberFormat = numberFormat;
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator2 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_0101;
				}
				continue;
				end_IL_0101:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
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
