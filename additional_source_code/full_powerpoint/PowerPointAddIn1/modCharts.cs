using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

[StandardModule]
public sealed class modCharts
{
	public static List<Axis> AxesList(this Chart cht)
	{
		List<Axis> list = new List<Axis>();
		using List<XlAxisType>.Enumerator enumerator = new List<XlAxisType>
		{
			XlAxisType.xlValue,
			XlAxisType.xlCategory,
			XlAxisType.xlSeriesAxis
		}.GetEnumerator();
		while (enumerator.MoveNext())
		{
			XlAxisType current = enumerator.Current;
			using List<XlAxisGroup>.Enumerator enumerator2 = new List<XlAxisGroup>
			{
				XlAxisGroup.xlPrimary,
				XlAxisGroup.xlSecondary
			}.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				XlAxisGroup current2 = enumerator2.Current;
				try
				{
					if (Conversions.ToBoolean(cht.get_HasAxis((object)current, (object)current2)))
					{
						Axis axis = (Axis)cht.Axes(current, current2);
						if (axis.Width != 0.0 && axis.Height != 0.0)
						{
							list.Add(axis);
						}
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				finally
				{
					Axis axis = null;
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
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			return list;
		}
	}
}
