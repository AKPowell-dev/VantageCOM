using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class AppoximateMatch
{
	internal static void A(ref List<ApproximateMatch> A, Severity B, Range C)
	{
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			string c = Conversions.ToString(NewLateBinding.LateGet(C, null, VH.A(1998), new object[0], null, null, null));
			AppoximateMatch.A(ref A, C, c, VH.A(2015));
			AppoximateMatch.A(ref A, C, c, VH.A(2030));
			AppoximateMatch.A(ref A, C, c);
			AppoximateMatch.B(ref A, C, c);
			AppoximateMatch.C(ref A, C, c);
			return;
		}
	}

	private static void A(ref List<ApproximateMatch> A, Range B, string C, string D)
	{
		bool result = false;
		if (!C.Contains(D))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = Helpers.A(C, D).GetEnumerator();
			while (enumerator.MoveNext())
			{
				string[] array = Helpers.A((Match)enumerator.Current);
				if (array.Length == 3)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					result = true;
				}
				else
				{
					string text = array[checked(array.Length - 1)];
					if (!bool.TryParse(text, out result))
					{
						if (int.TryParse(text, out var result2))
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
							result = result2 == 1;
						}
						else
						{
							Range range = Helpers.A(B.Worksheet, text);
							if (range != null)
							{
								Range range2 = range;
								if (range2.Cells.Count == 1)
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
									if (!bool.TryParse(Conversions.ToString(range2.Value2), out result))
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
										if (int.TryParse(Conversions.ToString(range2.Value2), out result2))
										{
											result = result2 == 1;
										}
									}
								}
								range2 = null;
								range = null;
							}
						}
					}
				}
				if (result)
				{
					A.Add(new ApproximateMatch(B, D));
					return;
				}
			}
			while (true)
			{
				switch (7)
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
	}

	private static void A(ref List<ApproximateMatch> A, Range B, string C)
	{
		bool flag = false;
		string text = VH.A(2045);
		if (!C.Contains(text))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				enumerator = Helpers.A(C, text).GetEnumerator();
				while (enumerator.MoveNext())
				{
					string[] array = Helpers.A((Match)enumerator.Current);
					if (array.Length == 2)
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
						flag = true;
					}
					else
					{
						string text2 = array[checked(array.Length - 1)];
						if (int.TryParse(text2, out var result))
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
							flag = result != 0;
						}
						else
						{
							Range range = Helpers.A(B.Worksheet, text2);
							if (range != null)
							{
								Range range2 = range;
								if (range2.Cells.Count == 1)
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
									if (int.TryParse(Conversions.ToString(range2.Value2), out result))
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
										flag = result != 0;
									}
								}
								range2 = null;
								range = null;
							}
						}
					}
					if (flag)
					{
						A.Add(new ApproximateMatch(B, text));
						return;
					}
				}
				while (true)
				{
					switch (7)
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

	private static void B(ref List<ApproximateMatch> A, Range B, string C)
	{
		bool flag = false;
		string text = VH.A(2056);
		if (!C.Contains(text))
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
				enumerator = Helpers.A(C, text).GetEnumerator();
				while (enumerator.MoveNext())
				{
					string[] array = Helpers.A((Match)enumerator.Current);
					switch (array.Length)
					{
					case 5:
					case 6:
					{
						if (int.TryParse(array[4], out var result))
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
							flag = AppoximateMatch.A(result);
						}
						else
						{
							flag = AppoximateMatch.A(B.Worksheet, array[4]);
						}
						break;
					}
					}
					if (!flag)
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
						A.Add(new ApproximateMatch(B, text));
						return;
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
	}

	private static void C(ref List<ApproximateMatch> A, Range B, string C)
	{
		bool flag = false;
		string text = VH.A(2071);
		if (!C.Contains(text))
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
				enumerator = Helpers.A(C, text).GetEnumerator();
				while (enumerator.MoveNext())
				{
					string[] array = Helpers.A((Match)enumerator.Current);
					if (array.Length != 2)
					{
						if (int.TryParse(array[2], out var result))
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							flag = AppoximateMatch.A(result);
						}
						else
						{
							flag = AppoximateMatch.A(B.Worksheet, array[2]);
						}
					}
					if (flag)
					{
						A.Add(new ApproximateMatch(B, text));
						return;
					}
				}
				while (true)
				{
					switch (7)
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
		}
	}

	private static bool A(Microsoft.Office.Interop.Excel.Worksheet A, string B)
	{
		Range range = Helpers.A(A, B);
		if (range != null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false) && int.TryParse(Conversions.ToString(range.Value2), out var result))
			{
				return AppoximateMatch.A(result);
			}
			range = null;
		}
		return false;
	}

	private static bool A(int A)
	{
		return A != 0;
	}
}
