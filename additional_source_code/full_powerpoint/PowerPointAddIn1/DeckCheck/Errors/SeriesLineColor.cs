using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SeriesLineColor : BaseColorError
{
	public SeriesLineColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, IMsoSeries series, Severity sev)
		: base(ErrorType.ColorPaletteBorder, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Series = series;
		((BaseError)this).Title = AH.A(24700);
		((BaseError)this).Subtitle = AH.A(24735);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		Microsoft.Office.Core.ColorFormat foreColor = ((BaseError)this).Series.Format.Line.ForeColor;
		int rGB = foreColor.RGB;
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		if (Charts.ImplsPoints(((BaseError)this).Series))
		{
			int num = 0;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ((IEnumerable)((BaseError)this).Series.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					num = checked(num + 1);
					try
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(objectValue, null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[0], null, null, null));
						if (!object.Equals(RuntimeHelpers.GetObjectValue(objectValue2), rGB))
						{
							dictionary.Add(num, Conversions.ToInteger(objectValue2));
						}
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (2)
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
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		foreColor.RGB = ColorTranslator.ToOle(color);
		foreach (KeyValuePair<int, int> item in dictionary)
		{
			try
			{
				NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(((BaseError)this).Series.Points(item.Key), null, AH.A(14028), new object[0], null, null, null), null, AH.A(14041), new object[0], null, null, null), null, AH.A(14050), new object[0], null, null, null), null, AH.A(14069), new object[1] { item.Value }, null, null, OptimisticSet: false, RValueBase: true);
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
		}
		foreColor = null;
	}
}
