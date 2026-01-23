using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Formulas
{
	internal static void A(WorksheetItem A, Range B)
	{
		Formulas.A(A, B, Formulas.A);
	}

	private static bool A(string A)
	{
		return A.Contains(VH.A(7827));
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Formulas.A(A, B, Formulas.B);
	}

	private static bool B(string A)
	{
		return A.Contains(VH.A(6144));
	}

	private static void A(WorksheetItem A, Range B, Func<string, bool> C)
	{
		Range A2 = null;
		B = RangeHelpers.A(B);
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				try
				{
					if (!C(range.Formula.ToString()))
					{
						continue;
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
						RangeHelpers.A(ref A2, range);
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
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_0073;
				}
				continue;
				end_IL_0073:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		S(A, A2);
		A2 = null;
	}

	internal static void C(WorksheetItem A, Range B)
	{
		B = RangeHelpers.A(B);
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Range A2 = null;
			string name = B.Worksheet.Name;
			Regex regex = new Regex(VH.A(103179));
			try
			{
				enumerator = B.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					try
					{
						if (regex.IsMatch(Formulas.A(Conversions.ToString(NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null)), name, range)))
						{
							RangeHelpers.A(ref A2, range);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_00bf;
					}
					continue;
					end_IL_00bf:
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
			S(A, A2);
			A2 = null;
			regex = null;
			return;
		}
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Range A2 = null;
		Range range = RangeHelpers.A(B);
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = range.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					if (!range2.Errors.get_Item((object)XlErrorChecks.xlEmptyCellReferences).Value)
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
					RangeHelpers.A(ref A2, range2);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0077;
					}
					continue;
					end_IL_0077:
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
			S(A, A2);
			A2 = null;
			range = null;
			return;
		}
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.C(B);
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			try
			{
				enumerator = range.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					if (range2.Columns.Count > 1)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						try
						{
							enumerator2 = range2.Rows.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Range a = (Range)enumerator2.Current;
								A.K(a);
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_008c;
								}
								continue;
								end_IL_008c:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (7)
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
					else
					{
						A.K(range2);
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_00ca;
					}
					continue;
					end_IL_00ca:
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
			range = null;
			return;
		}
	}

	internal static void F(WorksheetItem A, Range B)
	{
		Range range = null;
		try
		{
			range = RangeHelpers.C(B);
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
		IEnumerator enumerator = default(IEnumerator);
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
			int num = 1;
			try
			{
				enumerator = range.Areas.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Range a = (Range)enumerator.Current;
						A.K(a);
						if (num == 25)
						{
							A.B(Operators.SubtractObject(range.CountLarge, num).ToString() + VH.A(103254));
							break;
						}
						num = checked(num + 1);
						continue;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_00a5;
						}
						continue;
						end_IL_00a5:
						break;
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
						switch (2)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			range = null;
			return;
		}
	}

	internal static void G(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.H(B);
		if (range == null)
		{
			return;
		}
		Dictionary<XlErrorChecks, Range> dictionary = new Dictionary<XlErrorChecks, Range>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = range.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range2 = (Range)enumerator.Current;
				XlErrorChecks[] array = new XlErrorChecks[9]
				{
					XlErrorChecks.xlEvaluateToError,
					XlErrorChecks.xlEmptyCellReferences,
					XlErrorChecks.xlInconsistentFormula,
					XlErrorChecks.xlInconsistentListFormula,
					XlErrorChecks.xlListDataValidation,
					XlErrorChecks.xlNumberAsText,
					XlErrorChecks.xlOmittedCells,
					XlErrorChecks.xlTextDate,
					XlErrorChecks.xlUnlockedFormulaCells
				};
				foreach (XlErrorChecks xlErrorChecks in array)
				{
					if (!range2.Errors.get_Item((object)xlErrorChecks).Value)
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
					if (dictionary.ContainsKey(xlErrorChecks))
					{
						Range A2 = dictionary[xlErrorChecks];
						RangeHelpers.A(ref A2, range2);
						dictionary[xlErrorChecks] = A2;
						A2 = null;
					}
					else
					{
						dictionary.Add(xlErrorChecks, range2);
					}
					break;
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00db;
				}
				continue;
				end_IL_00db:
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
		using (Dictionary<XlErrorChecks, Range>.Enumerator enumerator2 = dictionary.GetEnumerator())
		{
			IEnumerator enumerator3 = default(IEnumerator);
			while (enumerator2.MoveNext())
			{
				KeyValuePair<XlErrorChecks, Range> current = enumerator2.Current;
				{
					enumerator3 = current.Value.Rows.GetEnumerator();
					try
					{
						while (enumerator3.MoveNext())
						{
							Range a = (Range)enumerator3.Current;
							A.A(a, current.Key);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_015e;
							}
							continue;
							end_IL_015e:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator3 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0189;
				}
				continue;
				end_IL_0189:
				break;
			}
		}
		range = null;
	}

	internal static void H(WorksheetItem A, Range B)
	{
		B = RangeHelpers.B(B);
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range a = (Range)enumerator.Current;
				A.E(a);
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
	}

	internal static void I(WorksheetItem A, Range B)
	{
		B = RangeHelpers.A(B, null, "", null);
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (1)
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
					Range a = (Range)enumerator.Current;
					A.E(a);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
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
	}

	internal static void J(WorksheetItem A, Range B)
	{
		Range A2 = null;
		B = RangeHelpers.A(B);
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				if (!ExcelAddIn1.Formulas.Helpers.ContainsPartialInput(range))
				{
					continue;
				}
				while (true)
				{
					switch (4)
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
				RangeHelpers.A(ref A2, range);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_005a;
				}
				continue;
				end_IL_005a:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		S(A, A2);
		A2 = null;
	}

	internal static void K(WorksheetItem A, Range B)
	{
		Range A2 = null;
		string input = Props.SearchForm.Input1;
		B = RangeHelpers.A(B);
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				if (!ExcelAddIn1.Formulas.Helpers.IsFunctionMatch(range, input))
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
				RangeHelpers.A(ref A2, range);
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
		S(A, A2);
		A2 = null;
	}

	internal static void L(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.A(B);
		Range A2 = null;
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Regex regex = new Regex(VH.A(103281));
			Regex regex2 = new Regex(VH.A(103362));
			string name = B.Worksheet.Name;
			enumerator = range.Cells.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					string text = Conversions.ToString(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null));
					if (text.EndsWith(VH.A(94843)))
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
						text = Formulas.A(text, name, range2);
						if (!regex.IsMatch(text))
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
						RangeHelpers.A(ref A2, range2);
						continue;
					}
					if (!text.Contains(VH.A(75498)))
					{
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					text = Formulas.A(text, name, range2);
					if (!regex2.IsMatch(text))
					{
						continue;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					RangeHelpers.A(ref A2, range2);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0149;
					}
					continue;
					end_IL_0149:
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
			R(A, A2);
			A2 = null;
			range = null;
			regex = null;
			regex2 = null;
			return;
		}
	}

	internal static void M(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.A(B);
		Range A2 = null;
		if (range == null)
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
			Regex regex = new Regex(VH.A(103479));
			string name = B.Worksheet.Name;
			enumerator = range.Cells.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					string text = Conversions.ToString(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null));
					if (!text.Contains(VH.A(75498)))
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
					text = Formulas.A(text, name, range2);
					if (!regex.IsMatch(text))
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
						break;
					}
					RangeHelpers.A(ref A2, range2);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_00e7;
					}
					continue;
					end_IL_00e7:
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
			R(A, A2);
			A2 = null;
			range = null;
			regex = null;
			return;
		}
	}

	internal static void N(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.A(B);
		Range A2 = null;
		if (range == null)
		{
			return;
		}
		string name = B.Worksheet.Name;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = range.Cells.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range2 = (Range)enumerator.Current;
				if (Operators.CompareString(Formulas.A(Conversions.ToString(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null)), name, range2), VH.A(54433), TextCompare: false) != 0)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				RangeHelpers.A(ref A2, range2);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_00aa;
				}
				continue;
				end_IL_00aa:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		R(A, A2);
		A2 = null;
		range = null;
	}

	internal static void O(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.A(B);
		Range A2 = null;
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Regex regex = new Regex(VH.A(103550));
			Regex regex2 = new Regex(VH.A(103623));
			string name = B.Worksheet.Name;
			try
			{
				enumerator = range.Cells.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					string text = Conversions.ToString(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null));
					if (!text.Contains(VH.A(75231)))
					{
						continue;
					}
					text = Formulas.A(text, name, range2);
					if (!regex.IsMatch(text))
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
						if (!regex2.IsMatch(text))
						{
							text = text.Replace(VH.A(39848), "").Replace(VH.A(39904), "");
							if (Operators.CompareString(text, VH.A(103696), TextCompare: false) == 0 || Operators.CompareString(text, VH.A(103739), TextCompare: false) == 0)
							{
								RangeHelpers.A(ref A2, range2);
							}
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
					}
					RangeHelpers.A(ref A2, range2);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_016a;
					}
					continue;
					end_IL_016a:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			R(A, A2);
			A2 = null;
			range = null;
			regex = null;
			regex2 = null;
			return;
		}
	}

	private static string A(string A, string B, Range C)
	{
		A = Regex.Replace(A, VH.A(103782), VH.A(48936));
		A = Regex.Replace(A, VH.A(103791), VH.A(48936));
		A = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(A, B);
		return Conversions.ToString(C.Application.ConvertFormula(A, XlReferenceStyle.xlA1, XlReferenceStyle.xlR1C1, XlReferenceType.xlRelative, C));
	}

	internal static void P(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.A(B);
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			enumerator = range.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Range instance = (Range)enumerator.Current;
					if (!Conversions.ToBoolean(NewLateBinding.LateGet(instance, null, VH.A(46494), new object[0], null, null, null)))
					{
						continue;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					A.H((Range)NewLateBinding.LateGet(instance, null, VH.A(103802), new object[0], null, null, null));
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_009d;
					}
					continue;
					end_IL_009d:
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
			range = null;
			return;
		}
	}

	internal static void Q(WorksheetItem A, Range B)
	{
		Range range = RangeHelpers.A(B);
		if (range == null)
		{
			return;
		}
		try
		{
			foreach (Range item in RangeHelpers.B(range))
			{
				A.I(item);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		range = null;
	}

	private static void R(WorksheetItem A, Range B)
	{
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
			enumerator = B.Rows.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Range a = (Range)enumerator.Current;
					A.E(a);
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
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
	}

	private static void S(WorksheetItem A, Range B)
	{
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			try
			{
				enumerator = B.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					if (range.Columns.Count > 1)
					{
						try
						{
							enumerator2 = range.Rows.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Range a = (Range)enumerator2.Current;
								A.E(a);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_007d;
								}
								continue;
								end_IL_007d:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (1)
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
					else
					{
						A.E(range);
					}
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return;
					}
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
		}
	}
}
