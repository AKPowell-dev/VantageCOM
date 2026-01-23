using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class PrivacySecurity
{
	[CompilerGenerated]
	internal sealed class P
	{
		public Range A;

		public Range B;

		public P(P A)
		{
			if (A == null)
			{
				return;
			}
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
				this.A = A.A;
				B = A.B;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			try
			{
				this.A = B.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
	}

	internal static void A(Analysis A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			List<Observation> observations = A.Observations;
			Range range = RangeHelpers.B(C);
			if (range != null)
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
				A.ActionStarted(VH.A(4587), range.Areas.Count);
				try
				{
					enumerator = range.Areas.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							Range rng = (Range)enumerator.Current;
							if (A.ItemCancelled())
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0089;
									}
									continue;
									end_IL_0089:
									break;
								}
								break;
							}
							observations.Add(new CellNote(B, rng));
							continue;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_00ad;
							}
							continue;
							end_IL_00ad:
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
				A.ActionEnded();
				range = null;
			}
			Range range2 = RangeHelpers.C(C);
			if (range2 != null)
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
				try
				{
					enumerator2 = range2.Areas.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Range rng2 = (Range)enumerator2.Current;
						observations.Add(new CellComment(B, rng2));
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_012d;
						}
						continue;
						end_IL_012d:
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
				range2 = null;
			}
			observations = null;
			return;
		}
	}

	internal static void A(Analysis A, ref ExcelAddIn1.Audit.Check.Observations.Raw.Observations B, Severity C, Range D)
	{
		P a = default(P);
		P CS_0024_003C_003E8__locals9 = new P(a);
		CS_0024_003C_003E8__locals9.B = D;
		if (C == Severity.Ignore)
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
			CS_0024_003C_003E8__locals9.A = null;
			global::A.A a2 = [SpecialName] () =>
			{
				try
				{
					CS_0024_003C_003E8__locals9.A = CS_0024_003C_003E8__locals9.B.SpecialCells(XlCellType.xlCellTypeConstants, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			};
			if (Operators.ConditionalCompareObjectGreater(CS_0024_003C_003E8__locals9.B.CountLarge, 40000, TextCompare: false))
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
				string a3 = VH.A(4642);
				global::A.A a4 = a2;
				object b;
				if (a4 != null)
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
					b = new Action(a4.Invoke);
				}
				else
				{
					b = null;
				}
				A.A(a3, (Action)b);
			}
			else
			{
				a2();
			}
			if (CS_0024_003C_003E8__locals9.A == null)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				Regex regex = new Regex(VH.A(4683));
				Regex regex2 = new Regex(VH.A(4700));
				Regex regex3 = new Regex(VH.A(4763));
				Regex regex4 = new Regex(VH.A(4804));
				Regex regex5 = new Regex(VH.A(4895));
				Regex regex6 = new Regex(VH.A(5067));
				Regex regex7 = new Regex(VH.A(5331));
				Regex regex8 = new Regex(VH.A(5398));
				Regex regex9 = new Regex(VH.A(5435), RegexOptions.IgnoreCase);
				Regex regex10 = new Regex(VH.A(5514));
				Regex regex11 = new Regex(VH.A(5585));
				new Regex(VH.A(5628));
				A.ActionStarted(VH.A(2512), Conversions.ToLong(CS_0024_003C_003E8__locals9.A.Cells.CountLarge));
				try
				{
					enumerator = CS_0024_003C_003E8__locals9.A.Cells.GetEnumerator();
					while (true)
					{
						List<ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData> sensitiveData;
						if (enumerator.MoveNext())
						{
							Range range = (Range)enumerator.Current;
							if (A.ItemCancelled())
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_01f6;
									}
									continue;
									end_IL_01f6:
									break;
								}
								break;
							}
							string input = Conversions.ToString(range.Text);
							string input2 = regex.Replace(input, "");
							sensitiveData = B.SensitiveData;
							if (regex2.IsMatch(input))
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
								sensitiveData.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData(range, VH.A(5711)));
							}
							else if (regex3.IsMatch(input))
							{
								sensitiveData.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData(range, VH.A(5718)));
							}
							else
							{
								if (!regex4.IsMatch(input))
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
									if (!regex7.IsMatch(input))
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
										if (!regex5.IsMatch(input2) && !regex6.IsMatch(input2))
										{
											if (!regex8.IsMatch(input2))
											{
												if (regex9.IsMatch(input))
												{
													sensitiveData.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData(range, VH.A(5762)));
												}
												else if (regex10.IsMatch(input))
												{
													sensitiveData.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData(range, VH.A(5789)));
												}
												else if (regex11.IsMatch(input))
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
													sensitiveData.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData(range, VH.A(5814)));
												}
												goto IL_0386;
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
										}
									}
								}
								sensitiveData.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData(range, VH.A(5725)));
							}
							goto IL_0386;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_0397;
							}
							continue;
							end_IL_0397:
							break;
						}
						break;
						IL_0386:
						sensitiveData = null;
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
				regex2 = null;
				regex3 = null;
				regex4 = null;
				regex5 = null;
				regex6 = null;
				regex7 = null;
				regex8 = null;
				regex9 = null;
				regex10 = null;
				regex11 = null;
				CS_0024_003C_003E8__locals9.A = null;
				return;
			}
		}
	}
}
