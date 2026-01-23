using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Format;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class MiscFormat
{
	internal static void A(WorksheetItem A, Range B)
	{
		Range range = null;
		Range A2 = null;
		try
		{
			range = B.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (range == null)
		{
			return;
		}
		Regex re = Footnotes.FootnoteRegex();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = range.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range2 = (Range)enumerator.Current;
				if (!Footnotes.CellContainsFootnote(range2, re))
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
				RangeHelpers.A(ref A2, range2);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0081;
				}
				continue;
				end_IL_0081:
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
		Helpers.B(A, A2);
		A2 = null;
		re = null;
		range = null;
	}

	internal static void B(WorksheetItem A, Range B)
	{
		using List<Range>.Enumerator enumerator = RangeHelpers.A(B, null, "").GetEnumerator();
		while (enumerator.MoveNext())
		{
			Range current = enumerator.Current;
			A.Q(current);
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			return;
		}
	}

	internal static void A(WorksheetItem A, object B)
	{
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					using List<Range>.Enumerator enumerator = RangeHelpers.A((Microsoft.Office.Interop.Excel.Worksheet)B, null, "").GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range current = enumerator.Current;
						A.O(current);
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				}
			}
		}
		Application application = MH.A.Application;
		using (List<Range>.Enumerator enumerator2 = RangeHelpers.A(((Range)B).Worksheet, null, "").GetEnumerator())
		{
			while (enumerator2.MoveNext())
			{
				Range current2 = enumerator2.Current;
				if (application.Intersect((Range)B, current2, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
				{
					A.O(current2);
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_0212;
				}
				continue;
				end_IL_0212:
				break;
			}
		}
		application = null;
	}
}
