using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class FormulaInterruption
{
	internal static void A(Analysis A, Severity B, Range C)
	{
		if (C == null)
		{
			return;
		}
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
			if (B == Severity.Ignore)
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				if (!FormulaInterruption.A(A, B, C))
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (FormulaInterruption.B(A, B, C))
					{
						FormulaInterruption.C(A, B, C);
					}
					return;
				}
			}
		}
	}

	private static bool A(Analysis A, Severity B, Range C)
	{
		if (true)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A.ActionStarted(VH.A(2658), Conversions.ToLong(C.Cells.CountLarge));
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = C.Cells.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						if (A.ItemCancelled())
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_006c;
								}
								continue;
								end_IL_006c:
								break;
							}
							break;
						}
						if (!range.Errors.get_Item((object)XlErrorChecks.xlInconsistentFormula).Value)
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
						A.Observations.Add(new ExcelAddIn1.Audit.Check.Observations.FormulaInterruption(B, range));
						continue;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_00bc;
						}
						continue;
						end_IL_00bc:
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
			A.ActionEnded();
			if (A.A())
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
		}
		else
		{
			A.ActionStarted(VH.A(2697), C.Areas.Count);
			{
				IEnumerator enumerator2 = C.Areas.GetEnumerator();
				try
				{
					while (true)
					{
						if (enumerator2.MoveNext())
						{
							Range range2 = (Range)enumerator2.Current;
							if (A.ItemCancelled())
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0159;
									}
									continue;
									end_IL_0159:
									break;
								}
								break;
							}
							if (Conversions.ToInteger(range2.Columns.CountLarge) > 1)
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
								{
									IEnumerator enumerator3 = range2.Rows.GetEnumerator();
									try
									{
										while (enumerator3.MoveNext())
										{
											Range range3 = (Range)enumerator3.Current;
											if (!FormulaInterruption.A(A.Observations, range3))
											{
												continue;
											}
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												break;
											}
											FormulaInterruption.B(A, B, range3);
										}
										while (true)
										{
											switch (3)
											{
											case 0:
												break;
											default:
												goto end_IL_01e0;
											}
											continue;
											end_IL_01e0:
											break;
										}
									}
									finally
									{
										IDisposable disposable2 = enumerator3 as IDisposable;
										if (disposable2 != null)
										{
											disposable2.Dispose();
										}
									}
								}
							}
							else
							{
								if (!FormulaInterruption.A(A.Observations, range2))
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
								FormulaInterruption.B(A, B, range2);
							}
							continue;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0231;
							}
							continue;
							end_IL_0231:
							break;
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
			A.ActionEnded();
			if (A.A())
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
		}
		return true;
	}

	private static bool A(List<Observation> A, Range B)
	{
		if (Conversions.ToLong(B.Cells.CountLarge) > 2)
		{
			Dictionary<string, Range> dictionary = new Dictionary<string, Range>();
			try
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = B.Cells.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						string key = range.FormulaR1C1.ToString();
						if (dictionary.ContainsKey(key))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						dictionary.Add(key, range);
						if (dictionary.Keys.Count <= 1)
						{
							continue;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							return true;
						}
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_009a;
						}
						continue;
						end_IL_009a:
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
			}
			finally
			{
				dictionary = null;
			}
		}
		return false;
	}

	private static void B(Analysis A, Severity B, Range C)
	{
		string objB = FormulaInterruption.A(C, 1L);
		int num = 1;
		long num2 = Conversions.ToLong(C.Cells.CountLarge);
		checked
		{
			A.ActionStarted(VH.A(2512), num2 - 1);
			long num3 = num2;
			long num4 = 2L;
			while (true)
			{
				if (num4 <= num3)
				{
					if (A.ItemCancelled())
					{
						break;
					}
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
						if (object.Equals(FormulaInterruption.A(C, num4), objB))
						{
							num++;
						}
						else
						{
							if (num > 1)
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
								if (object.Equals(FormulaInterruption.A(C, num4 + 1), objB))
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
									A.Observations.Add(new ExcelAddIn1.Audit.Check.Observations.FormulaInterruption(B, QB.A(C, num4)));
								}
							}
							objB = FormulaInterruption.A(C, num4);
							num = 1;
						}
						num4++;
						break;
					}
					continue;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				break;
			}
			A.ActionEnded();
		}
	}

	private static bool B(Analysis A, Severity B, Range C)
	{
		Application application = C.Application;
		Microsoft.Office.Interop.Excel.Worksheet worksheet = C.Worksheet;
		try
		{
			Dictionary<long, Range> dictionary = new Dictionary<long, Range>();
			A.ActionStarted(VH.A(2697), C.Areas.Count);
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = C.Areas.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						if (A.ItemCancelled())
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
								break;
							}
							break;
						}
						Range rows = range.Rows;
						int num = Conversions.ToInteger(rows.CountLarge);
						A.ActionStarted(VH.A(2758), num);
						int num2 = num;
						int num3 = 1;
						while (true)
						{
							if (num3 <= num2)
							{
								if (A.ItemCancelled())
								{
									break;
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									long num4 = QB.A(rows, num3).Row;
									Range range2 = application.Intersect(C, QB.A(worksheet.Rows, num4), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									if (range2 != null)
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
										if (!dictionary.ContainsKey(num4))
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												break;
											}
											dictionary.Add(num4, FormulaInterruption.A(worksheet, range2));
										}
									}
									num3 = checked(num3 + 1);
									break;
								}
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
							break;
						}
						A.ActionEnded();
						if (!A.A())
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
							return false;
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_02db;
						}
						continue;
						end_IL_02db:
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
			A.ActionEnded();
			if (A.A())
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
			FormulaInterruption.A(A, B: true, B, C, application, dictionary);
			return !A.A();
		}
		finally
		{
			Range rows = null;
			Dictionary<long, Range> dictionary = null;
			Range range2 = null;
			worksheet = null;
			application = null;
		}
	}

	private static bool C(Analysis A, Severity B, Range C)
	{
		Application application = C.Application;
		Microsoft.Office.Interop.Excel.Worksheet worksheet = C.Worksheet;
		try
		{
			Dictionary<long, Range> dictionary = new Dictionary<long, Range>();
			A.ActionStarted(VH.A(2697), C.Areas.Count);
			IEnumerator enumerator = C.Areas.GetEnumerator();
			try
			{
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						if (A.ItemCancelled())
						{
							while (true)
							{
								switch (7)
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
							break;
						}
						Range columns = range.Columns;
						int num = Conversions.ToInteger(columns.CountLarge);
						A.ActionStarted(VH.A(2789), num);
						int num2 = num;
						int num3 = 1;
						while (true)
						{
							if (num3 <= num2)
							{
								if (A.ItemCancelled())
								{
									break;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									long num4 = QB.A(columns, num3).Column;
									Range range2 = application.Intersect(C, QB.A(worksheet.Columns, num4), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									if (range2 != null)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											break;
										}
										if (!dictionary.ContainsKey(num4))
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
											dictionary.Add(num4, FormulaInterruption.A(worksheet, range2));
										}
									}
									num3 = checked(num3 + 1);
									break;
								}
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
							break;
						}
						A.ActionEnded();
						if (A.A())
						{
							return false;
						}
						continue;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_02b4;
						}
						continue;
						end_IL_02b4:
						break;
					}
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
			A.ActionEnded();
			if (A.A())
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
			FormulaInterruption.A(A, B: false, B, C, application, dictionary);
			return !A.A();
		}
		finally
		{
			Dictionary<long, Range> dictionary = null;
			Range range2 = null;
			worksheet = null;
			application = null;
		}
	}

	private static Range A(Microsoft.Office.Interop.Excel.Worksheet A, Range B)
	{
		return ((_Worksheet)A).get_Range((object)B.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Replace(VH.A(2378), VH.A(2826)), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	private static void A(Analysis A, bool B, Severity C, Range D, Application E, Dictionary<long, Range> F)
	{
		string format = VH.A(2829);
		string arg;
		if (!B)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			arg = VH.A(2862);
		}
		else
		{
			arg = VH.A(2877);
		}
		A.ActionStarted(string.Format(format, arg), F.LongCount());
		checked
		{
			using (Dictionary<long, Range>.Enumerator enumerator = F.GetEnumerator())
			{
				do
				{
					if (enumerator.MoveNext())
					{
						KeyValuePair<long, Range> current = enumerator.Current;
						if (A.ItemCancelled())
						{
							break;
						}
						Range value = current.Value;
						string objB = FormulaInterruption.A(value, 1L);
						int num = 1;
						long num2 = Conversions.ToLong(value.Cells.CountLarge);
						string actionDesc;
						if (!B)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
							actionDesc = VH.A(2886);
						}
						else
						{
							actionDesc = VH.A(2933);
						}
						A.ActionStarted(actionDesc, num2 - 2);
						long num3 = num2 - 1;
						long num4 = 2L;
						while (true)
						{
							if (num4 <= num3)
							{
								if (A.ItemCancelled())
								{
									break;
								}
								if (E.Intersect(QB.A(value, num4), D, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
									if (object.Equals(FormulaInterruption.A(value, num4), objB))
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											break;
										}
										num++;
									}
									else
									{
										objB = FormulaInterruption.A(value, num4);
										num = 1;
									}
								}
								else
								{
									if (num > 1 && FormulaInterruption.A(value, num4 + 1) && object.Equals(FormulaInterruption.A(value, num4 + 1), objB))
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
										A.Observations.Add(new ExcelAddIn1.Audit.Check.Observations.FormulaInterruption(C, QB.A(value, num4)));
									}
									objB = "";
									num = 0;
								}
								num4++;
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
							break;
						}
						A.ActionEnded();
						value = null;
						continue;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_031f;
						}
						continue;
						end_IL_031f:
						break;
					}
					break;
				}
				while (!A.A());
			}
			A.ActionEnded();
		}
	}

	private static string A(Range A, long B)
	{
		return QB.A(A, B).FormulaR1C1.ToString();
	}

	private static bool A(Range A, long B)
	{
		return object.Equals(RuntimeHelpers.GetObjectValue(QB.A(A, B).HasFormula), true);
	}
}
