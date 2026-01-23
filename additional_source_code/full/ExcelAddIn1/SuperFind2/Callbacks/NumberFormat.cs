using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class NumberFormat
{
	internal static void A(WorksheetItem A, Range B)
	{
		NumberFormat.A(A, B, NumberFormat.A);
	}

	private static bool A(string A, string B)
	{
		return Operators.CompareString(A, B, TextCompare: false) == 0;
	}

	internal static void B(WorksheetItem A, Range B)
	{
		NumberFormat.A(A, B, NumberFormat.B);
	}

	private static bool B(string A, string B)
	{
		return A.Contains(B);
	}

	private static void A(WorksheetItem A, Range B, Func<string, string, bool> C)
	{
		string input = Props.SearchForm.Input1;
		Range A2 = null;
		if (!Props.SearchForm.LookInEmptyCells)
		{
			B = RangeHelpers.H(B);
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = B.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					try
					{
						if (!C(range.NumberFormat.ToString(), input))
						{
							continue;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							RangeHelpers.A(ref A2, range);
							break;
						}
						continue;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
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
			Helpers.B(A, A2);
			A2 = null;
			return;
		}
	}

	internal static void C(WorksheetItem A, Range B)
	{
		List<string> list = new List<string>();
		Range A2 = null;
		if (!(B.Application.Selection is Range))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((Range)B.Application.Selection).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				list.Add(range.NumberFormat.ToString());
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		list = list.Distinct().ToList();
		if (!Props.SearchForm.LookInEmptyCells)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			B = RangeHelpers.H(B);
		}
		if (B != null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			foreach (Range item in B)
			{
				try
				{
					if (!list.Contains(item.NumberFormat.ToString()))
					{
						continue;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						RangeHelpers.A(ref A2, item);
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			Helpers.B(A, A2);
			A2 = null;
		}
		list = null;
	}
}
