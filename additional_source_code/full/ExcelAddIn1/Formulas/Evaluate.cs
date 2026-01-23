using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.UndoRedo;
using ExcelAddIn1.Workbook;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Evaluate
{
	[CompilerGenerated]
	internal sealed class BG
	{
		public string A;

		public Func<string> A;

		public BG(BG A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal string A()
		{
			return this.A;
		}

		[SpecialName]
		internal string B()
		{
			return this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class CG
	{
		public Range A;

		public CG(CG A)
		{
			if (A == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal string A()
		{
			return this.A.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
	}

	[CompilerGenerated]
	internal sealed class DG
	{
		public double A;

		public double? A;

		[SpecialName]
		internal bool A(double A)
		{
			return this.A * A < this.A * this.A.Value;
		}
	}

	[CompilerGenerated]
	internal sealed class EG
	{
		public string A;

		public Func<string> A;

		public EG(EG A)
		{
			if (A == null)
			{
				return;
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal string A()
		{
			return this.A;
		}

		[SpecialName]
		internal string B()
		{
			return this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class FG
	{
		public Range A;

		public FG(FG A)
		{
			if (A == null)
			{
				return;
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal string A()
		{
			return this.A.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
	}

	private static long m_A;

	public static void SimplifyFormula()
	{
		if (!A())
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			Range range = null;
			int E = 0;
			if (application.Selection is Range)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
					{
						Range range2 = JH.A((Range)application.Selection);
						if (Conversions.ToBoolean(application.ActiveCell.HasArray))
						{
							Forms.WarningMessage(VH.A(158117));
							application = null;
						}
						else
						{
							range = B(range2);
							if (range != null)
							{
								List<Range> D = new List<Range>();
								application.ScreenUpdating = false;
								application.EnableEvents = false;
								MySettings settings = K.Settings;
								bool simplifyIndirect = settings.SimplifyIndirect;
								bool simplifyChoose = settings.SimplifyChoose;
								bool simplifyOffset = settings.SimplifyOffset;
								bool simplifyHlookup = settings.SimplifyHlookup;
								bool simplifyVlookup = settings.SimplifyVlookup;
								bool simplifyXlookup = settings.SimplifyXlookup;
								bool simplifyIndexMatch = settings.SimplifyIndexMatch;
								bool simplifyIf = settings.SimplifyIf;
								bool simplifyMin = settings.SimplifyMin;
								bool simplifyMax = settings.SimplifyMax;
								bool simplifySumIf = settings.SimplifySumIf;
								bool simplifySumIfs = settings.SimplifySumIfs;
								_ = null;
								bool flag = JH.A(range2);
								try
								{
									enumerator = range.GetEnumerator();
									try
									{
										while (enumerator.MoveNext())
										{
											Range range3 = (Range)enumerator.Current;
											if (Conversions.ToBoolean(Operators.NotObject(range3.HasArray)))
											{
												string text = range3.Formula.ToString();
												string strFormula = text;
												object objectValue = RuntimeHelpers.GetObjectValue(range3.Value2);
												bool blnError = false;
												if (simplifyVlookup)
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
													SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(2015), EvaluateLookup);
												}
												if (simplifyHlookup && !blnError)
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
													SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(2030), EvaluateLookup);
												}
												if (simplifyXlookup)
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
													if (!blnError)
													{
														SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(2056), EvaluateXLookup);
													}
												}
												if (simplifyIndirect && !blnError)
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
													SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(4457), B);
												}
												if (simplifyChoose)
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
													if (!blnError)
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
														SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(158202), C);
													}
												}
												if (simplifyOffset)
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
													if (!blnError)
													{
														SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(4444), Evaluate.D);
													}
												}
												if (simplifyMin && !blnError)
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
													SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(49095), EvaluateMin);
												}
												if (simplifyMax && !blnError)
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
													SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(49102), EvaluateMax);
												}
												if (simplifySumIfs)
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
													if (!blnError)
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
														SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(49109), EvaluateSumIfs);
													}
												}
												if (simplifySumIf && !blnError)
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
													SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(4533), EvaluateSumIf);
												}
												if (simplifyIf)
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
													if (!blnError)
													{
														SimplifyFunction(range3, ref strFormula, ref blnError, VH.A(3794), A);
													}
												}
												if (simplifyIndexMatch && !blnError)
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
													A(ref strFormula, range3, range3.Worksheet, ref D);
												}
												if (!blnError)
												{
													A(range3, RuntimeHelpers.GetObjectValue(objectValue), strFormula, text, ref E, ref D);
												}
												else
												{
													D.Add(range3);
												}
											}
										}
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												goto end_IL_045a;
											}
											continue;
											end_IL_045a:
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
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									clsReporting.LogException(ex2);
									ProjectData.ClearProjectError();
								}
								if (flag)
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
									JH.A(range2, VH.A(158215));
								}
								Common.RefreshLiveVisualizations(range2);
								application.ScreenUpdating = true;
								application.EnableEvents = true;
								D = D.Distinct().ToList();
								int count = D.Count;
								if (range.Count == 1)
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
									if (E != 1)
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
										if (count == 1)
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
											Forms.WarningMessage(VH.A(158248));
										}
										else
										{
											Forms.InfoMessage(VH.A(158313));
										}
									}
								}
								else if (E > 0)
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
									if (count == 0)
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
										Forms.SuccessMessage(VH.A(158386) + E + VH.A(158409));
									}
									else if (A(VH.A(158386) + E + VH.A(158434) + count + VH.A(158507)))
									{
										Miscellaneous.GoToDialog(D);
									}
								}
								else if (count == 0)
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
									Forms.InfoMessage(VH.A(158606));
								}
								else if (A(VH.A(158687) + count + VH.A(158726)))
								{
									Miscellaneous.GoToDialog(D);
								}
								clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(158215));
								D = null;
								range = null;
							}
							else
							{
								Forms.InfoMessage(VH.A(158821));
							}
							application = null;
							range2 = null;
						}
						return;
					}
					}
				}
			}
			Forms.WarningMessage(VH.A(158034));
			application = null;
			return;
		}
	}

	public static void SimplifyFunction(Range rng, ref string strFormula, ref bool blnError, string strFunction, Func<string, string, string> f)
	{
		SimplifyFunction(rng.Worksheet, ref strFormula, ref blnError, strFunction, f);
	}

	public static void SimplifyFunction(Worksheet ws, ref string strFormula, ref bool blnError, string strFunction, Func<string, string, string> f)
	{
		if (B(strFormula, strFunction) <= 0)
		{
			return;
		}
		checked
		{
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
				strFormula = Strings.Right(strFormula, Strings.Len(strFormula) - 1);
				for (int num = B(strFormula, strFunction); num > 0; num = B(strFormula, strFunction))
				{
					string text = D(strFormula, num);
					string replacement;
					try
					{
						replacement = f(text, strFunction);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						blnError = true;
						ProjectData.ClearProjectError();
						return;
					}
					strFormula = Strings.Replace(strFormula, text, replacement, 1, 1);
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (Operators.CompareString(strFunction, VH.A(158202), TextCompare: false) != 0)
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
						if (Strings.InStr(1, strFormula, VH.A(44119)) > 0)
						{
							blnError = true;
							return;
						}
					}
					string text2 = VH.A(48936);
					if (Versioned.IsNumeric(strFormula))
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
						text2 = "";
					}
					strFormula = text2 + strFormula;
					string left = Strings.Mid(strFormula, 2, 1);
					string left2 = Strings.Right(strFormula, 1);
					string text3 = Strings.Replace(strFormula, VH.A(39830), "");
					if (Operators.CompareString(left, VH.A(39830), TextCompare: false) != 0)
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
						if (Operators.CompareString(left2, VH.A(39830), TextCompare: false) != 0)
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
							if (strFormula.Length - text3.Length == 2)
							{
								try
								{
									strFormula = Conversions.ToString(ws.Evaluate(strFormula));
									return;
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									blnError = true;
									ProjectData.ClearProjectError();
									return;
								}
							}
							return;
						}
					}
				}
			}
		}
	}

	public static void IfFunction()
	{
		if (!A())
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
			A(VH.A(3794), A);
			return;
		}
	}

	private static string A(string A, string B)
	{
		object[] array = new object[1];
		long num = 0L;
		checked
		{
			Evaluate.m_A = B.Length + 2;
			while (Evaluate.m_A < A.Length)
			{
				array = (object[])Utils.CopyArray(array, new object[(int)num + 1]);
				array[(int)num] = C(A, (int)Evaluate.m_A);
				num++;
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
				string name = Conversions.ToString(array[0]);
				object objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Evaluate(name));
				if (objectValue is bool)
				{
					if (Conversions.ToBoolean(objectValue))
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								return Conversions.ToString(array[1]);
							}
						}
					}
					return Conversions.ToString(array[2]);
				}
				return A;
			}
		}
	}

	public static void Min()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(VH.A(49095), EvaluateMin);
			return;
		}
	}

	public static string EvaluateMin(string strFormula, string strFunction)
	{
		EG a = default(EG);
		EG CS_0024_003C_003E8__locals9 = new EG(a);
		new List<string>();
		List<double> list = new List<double>();
		string D = strFormula;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			Evaluate.m_A = strFunction.Length + 2;
			IEnumerator enumerator = default(IEnumerator);
			FG fG = default(FG);
			while (Evaluate.m_A < strFormula.Length)
			{
				CS_0024_003C_003E8__locals9.A = C(strFormula, (int)Evaluate.m_A);
				if (application.Evaluate(CS_0024_003C_003E8__locals9.A) is Range)
				{
					Range range = (Range)application.Evaluate(CS_0024_003C_003E8__locals9.A);
					if (!Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
					{
						{
							enumerator = range.GetEnumerator();
							try
							{
								while (enumerator.MoveNext())
								{
									fG = new FG(fG);
									fG.A = (Range)enumerator.Current;
									A(list, RuntimeHelpers.GetObjectValue(fG.A.Value2), C: true, ref D, fG.A);
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0148;
									}
									continue;
									end_IL_0148:
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
						}
					}
					else
					{
						List<double> a2 = list;
						object objectValue = RuntimeHelpers.GetObjectValue(range.Value2);
						Func<string> e;
						if (CS_0024_003C_003E8__locals9.A != null)
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							e = CS_0024_003C_003E8__locals9.A;
						}
						else
						{
							e = (CS_0024_003C_003E8__locals9.A = [SpecialName] () => CS_0024_003C_003E8__locals9.A);
						}
						A(a2, objectValue, C: true, ref D, e);
					}
					range = null;
				}
				else
				{
					A(list, RuntimeHelpers.GetObjectValue(application.Evaluate(CS_0024_003C_003E8__locals9.A)), C: true, ref D, [SpecialName] () => CS_0024_003C_003E8__locals9.A);
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				list = null;
				application = null;
				return D;
			}
		}
	}

	public static void Max()
	{
		if (!A())
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
			A(VH.A(49102), EvaluateMax);
			return;
		}
	}

	public static string EvaluateMax(string strFormula, string strFunction)
	{
		BG a = default(BG);
		BG CS_0024_003C_003E8__locals6 = new BG(a);
		new List<string>();
		List<double> a2 = new List<double>();
		string D = strFormula;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			Evaluate.m_A = strFunction.Length + 2;
			IEnumerator enumerator = default(IEnumerator);
			CG cG = default(CG);
			while (Evaluate.m_A < strFormula.Length)
			{
				CS_0024_003C_003E8__locals6.A = C(strFormula, (int)Evaluate.m_A);
				if (application.Evaluate(CS_0024_003C_003E8__locals6.A) is Range)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Range range = (Range)application.Evaluate(CS_0024_003C_003E8__locals6.A);
					if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
						A(a2, RuntimeHelpers.GetObjectValue(range.Value2), C: false, ref D, [SpecialName] () => CS_0024_003C_003E8__locals6.A);
					}
					else
					{
						try
						{
							enumerator = range.GetEnumerator();
							while (enumerator.MoveNext())
							{
								cG = new CG(cG);
								cG.A = (Range)enumerator.Current;
								A(a2, RuntimeHelpers.GetObjectValue(cG.A.Value2), C: false, ref D, cG.A);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0158;
								}
								continue;
								end_IL_0158:
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
					}
					range = null;
				}
				else
				{
					A(a2, RuntimeHelpers.GetObjectValue(application.Evaluate(CS_0024_003C_003E8__locals6.A)), C: false, ref D, [SpecialName] () => CS_0024_003C_003E8__locals6.A);
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				a2 = null;
				application = null;
				return D;
			}
		}
	}

	public static void Indirect()
	{
		if (!A())
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
			A(VH.A(4457), B);
			return;
		}
	}

	private static string B(string A, string B)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		object[] array = default(object[]);
		int num = default(int);
		int num3 = default(int);
		long num5 = default(long);
		long a = default(long);
		string text = default(string);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		string result = default(string);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						array = new object[1];
						goto IL_000a;
					case 504:
						{
							num = num2;
							switch (num3)
							{
							case 1:
								break;
							default:
								goto end_IL_0000;
							}
							int num4 = unchecked(num + 1);
							num = 0;
							switch (num4)
							{
							case 1:
								break;
							case 2:
								goto IL_000a;
							case 3:
								goto IL_0010;
							case 4:
								goto IL_001d;
							case 7:
								goto IL_0028;
							case 8:
								goto IL_0044;
							case 9:
								goto IL_0058;
							case 5:
							case 6:
							case 10:
								goto IL_0062;
							case 11:
								goto IL_0085;
							case 12:
								goto IL_008c;
							case 13:
								goto IL_0098;
							case 14:
								goto IL_00ab;
							case 15:
								goto IL_00c9;
							case 17:
								goto IL_0104;
							case 16:
							case 18:
								goto IL_0124;
							case 20:
								goto IL_0129;
							case 19:
							case 21:
								goto IL_013a;
							case 22:
								goto IL_0141;
							case 23:
								goto IL_015d;
							case 24:
								goto IL_0164;
							case 25:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 26:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_015d:
						ProjectData.ClearProjectError();
						num3 = 0;
						goto IL_0164;
						IL_000a:
						num2 = 2;
						num5 = 0L;
						goto IL_0010;
						IL_0010:
						num2 = 3;
						a = Strings.Len(B) + 2;
						goto IL_001d;
						IL_001d:
						num2 = 4;
						Evaluate.m_A = a;
						goto IL_0062;
						IL_0062:
						num2 = 6;
						if (Evaluate.m_A < Strings.Len(A))
						{
							goto IL_0028;
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
						goto IL_0085;
						IL_0164:
						num2 = 24;
						text = VH.A(44119);
						break;
						IL_0028:
						num2 = 7;
						array = (object[])Utils.CopyArray(array, new object[(int)num5 + 1]);
						goto IL_0044;
						IL_0044:
						num2 = 8;
						array[(int)num5] = C(A, (int)Evaluate.m_A);
						goto IL_0058;
						IL_0085:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_008c;
						IL_008c:
						num2 = 12;
						if (num5 == 1)
						{
							goto IL_0098;
						}
						goto IL_0129;
						IL_0098:
						num2 = 13;
						application = MH.A.Application;
						goto IL_00ab;
						IL_00ab:
						num2 = 14;
						if (application.Evaluate(RuntimeHelpers.GetObjectValue(array[0])) is Range)
						{
							goto IL_00c9;
						}
						goto IL_0104;
						IL_00c9:
						num2 = 15;
						text = Conversions.ToString(NewLateBinding.LateGet(application.Evaluate(RuntimeHelpers.GetObjectValue(array[0])), null, VH.A(41636), new object[0], null, null, null));
						goto IL_0124;
						IL_0104:
						num2 = 17;
						text = Conversions.ToString(application.Evaluate(RuntimeHelpers.GetObjectValue(array[0])));
						goto IL_0124;
						IL_0124:
						application = null;
						goto IL_013a;
						IL_0129:
						num2 = 20;
						text = VH.A(44119);
						goto IL_013a;
						IL_013a:
						ProjectData.ClearProjectError();
						num3 = 0;
						goto IL_0141;
						IL_0141:
						num2 = 22;
						if (Information.Err().Number <= 0)
						{
							break;
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
						goto IL_015d;
						IL_0058:
						num2 = 9;
						num5++;
						goto IL_0062;
						end_IL_0000_2:
						break;
					}
					num2 = 25;
					result = text;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 504;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void Choose()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(VH.A(158202), C);
			return;
		}
	}

	private static string C(string A, string B)
	{
		object[] array = new object[1];
		long num = 0L;
		checked
		{
			Evaluate.m_A = B.Length + 2;
			while (Evaluate.m_A < A.Length)
			{
				array = (object[])Utils.CopyArray(array, new object[(int)num + 1]);
				array[(int)num] = C(A, (int)Evaluate.m_A);
				num++;
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
				long num2 = Evaluate.A(RuntimeHelpers.GetObjectValue(array[0]));
				return Conversions.ToString(array[(int)num2]);
			}
		}
	}

	public static void Offset()
	{
		if (!A())
		{
			return;
		}
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
			A(VH.A(4444), D);
			return;
		}
	}

	private static string D(string A, string B)
	{
		object[] array = new object[1];
		long num = 0L;
		checked
		{
			Evaluate.m_A = B.Length + 2;
			while (Evaluate.m_A < A.Length)
			{
				array = (object[])Utils.CopyArray(array, new object[(int)num + 1]);
				array[(int)num] = C(A, (int)Evaluate.m_A);
				num++;
			}
			string[] array2 = default(string[]);
			string a = default(string);
			string text = default(string);
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
					if (Strings.InStr(1, Conversions.ToString(array[0]), VH.A(7827)) > 0)
					{
						array2 = Strings.Split(Conversions.ToString(array[0]), VH.A(7827));
						a = array2[1];
					}
					else
					{
						a = Conversions.ToString(array[0]);
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				long b;
				string result;
				try
				{
					string left = array[1].ToString();
					if (Operators.CompareString(left, VH.A(96037), TextCompare: false) != 0)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							if (Operators.CompareString(left, "", TextCompare: false) != 0 && Operators.CompareString(left, VH.A(64024), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									b = Evaluate.A(RuntimeHelpers.GetObjectValue(array[1]));
									break;
								}
							}
							else
							{
								b = 0L;
							}
							break;
						}
					}
					else
					{
						b = 1L;
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					result = VH.A(158892);
					ProjectData.ClearProjectError();
					goto IL_0422;
				}
				long c;
				try
				{
					string left2 = array[2].ToString();
					if (Operators.CompareString(left2, VH.A(96037), TextCompare: false) != 0)
					{
						if (Operators.CompareString(left2, "", TextCompare: false) == 0)
						{
							goto IL_01d0;
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
						if (Operators.CompareString(left2, VH.A(64024), TextCompare: false) == 0)
						{
							goto IL_01d0;
						}
						c = Evaluate.A(RuntimeHelpers.GetObjectValue(array[2]));
					}
					else
					{
						c = 1L;
					}
					goto end_IL_0175;
					IL_01d0:
					c = 0L;
					end_IL_0175:;
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					result = VH.A(158892);
					ProjectData.ClearProjectError();
					goto IL_0422;
				}
				long d;
				try
				{
					if (num > 3)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							string left3 = array[3].ToString();
							if (Operators.CompareString(left3, VH.A(96037), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									if (Operators.CompareString(left3, "", TextCompare: false) != 0)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											if (Operators.CompareString(left3, VH.A(64024), TextCompare: false) != 0)
											{
												while (true)
												{
													switch (2)
													{
													case 0:
														continue;
													}
													d = Evaluate.A(RuntimeHelpers.GetObjectValue(array[3]));
													break;
												}
												break;
											}
											result = VH.A(44108);
											goto IL_0422;
										}
									}
									else
									{
										d = 0L;
									}
									break;
								}
							}
							else
							{
								d = 1L;
							}
							break;
						}
					}
					else
					{
						d = 0L;
					}
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					result = VH.A(44108);
					ProjectData.ClearProjectError();
					goto IL_0422;
				}
				long e;
				try
				{
					if (num > 4)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							string left4 = array[4].ToString();
							if (Operators.CompareString(left4, VH.A(96037), TextCompare: false) != 0)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									if (Operators.CompareString(left4, "", TextCompare: false) != 0)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											if (Operators.CompareString(left4, VH.A(64024), TextCompare: false) == 0)
											{
												result = VH.A(44108);
												goto IL_0422;
											}
											e = Evaluate.A(RuntimeHelpers.GetObjectValue(array[4]));
											break;
										}
									}
									else
									{
										e = 0L;
									}
									break;
								}
							}
							else
							{
								e = 1L;
							}
							break;
						}
					}
					else
					{
						e = 0L;
					}
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					result = VH.A(158892);
					ProjectData.ClearProjectError();
					goto IL_0422;
				}
				try
				{
					text = Strings.Replace(Evaluate.A(a, b, c, d, e), VH.A(41262), "");
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
				try
				{
					result = ((Information.UBound(array2) != 1) ? text : (array2[0] + VH.A(7827) + text));
				}
				catch (Exception ex13)
				{
					ProjectData.SetProjectError(ex13);
					Exception ex14 = ex13;
					result = text;
					ProjectData.ClearProjectError();
				}
				goto IL_0422;
				IL_0422:
				return result;
			}
		}
	}

	private static string A(string A, long B, long C, long D, long E)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		string result;
		int num = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
				{
					ProjectData.ClearProjectError();
					num2 = 0;
					object activeSheet = MH.A.Application.ActiveSheet;
					string memberName = VH.A(41315);
					object[] obj = new object[1] { A };
					object[] array = obj;
					bool[] obj2 = new bool[1] { true };
					bool[] array2 = obj2;
					object instance = NewLateBinding.LateGet(activeSheet, null, memberName, obj, null, null, obj2);
					if (array2[0])
					{
						A = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
					}
					object[] array3;
					bool[] array4;
					object obj3 = NewLateBinding.LateGet(instance, null, VH.A(60565), array3 = new object[2] { B, C }, null, null, array4 = new bool[2] { true, true });
					if (array4[0])
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
						B = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(long));
					}
					if (array4[1])
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
						C = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
					}
					Range range = (Range)obj3;
					if (!(D == 0 && E == 0))
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
						if (D == 0L)
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
							range = range.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)E);
						}
						else if (E == 0L)
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
							range = range.get_Resize((object)D, RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						else
						{
							range = range.get_Resize((object)D, (object)E);
						}
					}
					if (Information.Err().Number > 0)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								ProjectData.ClearProjectError();
								num2 = 0;
								Information.Err().Number = 0;
								result = VH.A(158892);
								goto end_IL_0000;
							}
						}
					}
					result = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto end_IL_0000;
				}
				case 538:
					num = -1;
					switch (num2)
					{
					}
					break;
				}
			}
			catch (object obj4) when (obj4 is Exception && num2 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj4);
				try0000_dispatch = 538;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void XLookup()
	{
		if (!A())
		{
			return;
		}
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
			A(VH.A(2056), EvaluateXLookup);
			return;
		}
	}

	public static string EvaluateXLookup(string strFormula, string strFunction)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		object[] array = default(object[]);
		int num = default(int);
		int num3 = default(int);
		List<string> list = default(List<string>);
		string item = default(string);
		int num5 = default(int);
		long a = default(long);
		string item2 = default(string);
		string text = default(string);
		string item3 = default(string);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		string[] array2 = default(string[]);
		Name name = default(Name);
		string item4 = default(string);
		object objectValue = default(object);
		string[] array3 = default(string[]);
		string[] array4 = default(string[]);
		Microsoft.Office.Interop.Excel.Application application2 = default(Microsoft.Office.Interop.Excel.Application);
		object obj = default(object);
		string item5 = default(string);
		string result = default(string);
		object obj2 = default(object);
		object obj3 = default(object);
		object obj4 = default(object);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						array = new object[1];
						goto IL_000a;
					case 2902:
						{
							num = num2;
							switch (num3)
							{
							case 1:
								break;
							default:
								goto end_IL_0000;
							}
							int num4 = unchecked(num + 1);
							num = 0;
							switch (num4)
							{
							case 1:
								break;
							case 2:
								goto IL_000a;
							case 3:
								goto IL_000f;
							case 4:
								goto IL_001e;
							case 7:
								goto IL_0029;
							case 8:
								goto IL_0044;
							case 9:
								goto IL_0059;
							case 5:
							case 6:
							case 10:
								goto IL_0062;
							case 11:
								goto IL_0087;
							case 12:
								goto IL_008e;
							case 13:
								goto IL_00a1;
							case 14:
								goto IL_00ad;
							case 15:
								goto IL_00b8;
							case 16:
								goto IL_00c4;
							case 17:
								goto IL_00fa;
							case 18:
								goto IL_0117;
							case 20:
								goto IL_014a;
							case 19:
							case 21:
								goto IL_0158;
							case 22:
								goto IL_0189;
							case 23:
								goto IL_01ac;
							case 25:
								goto IL_01eb;
							case 24:
							case 26:
								goto IL_0209;
							case 28:
								goto IL_0216;
							case 27:
							case 29:
								goto IL_0228;
							case 30:
								goto IL_022f;
							case 31:
								goto IL_023a;
							case 32:
								goto IL_0246;
							case 33:
								goto IL_027c;
							case 34:
								goto IL_0297;
							case 36:
								goto IL_02d0;
							case 35:
							case 37:
								goto IL_02e0;
							case 38:
								goto IL_0309;
							case 39:
								goto IL_032c;
							case 41:
								goto IL_036b;
							case 40:
							case 42:
								goto IL_0387;
							case 44:
								goto IL_0394;
							case 43:
							case 45:
								goto IL_03a4;
							case 46:
								goto IL_03aa;
							case 47:
								goto IL_03b4;
							case 48:
								goto IL_03be;
							case 49:
								goto IL_03c8;
							case 50:
								goto IL_03cf;
							case 51:
								goto IL_03ed;
							case 52:
								goto IL_03f9;
							case 53:
								goto IL_043b;
							case 55:
								goto IL_0443;
							case 56:
								goto IL_0459;
							case 58:
								goto IL_048f;
							case 54:
							case 57:
							case 59:
								goto IL_049d;
							case 60:
								goto IL_04e4;
							case 61:
								goto IL_052f;
							case 62:
								goto IL_0544;
							case 63:
								goto IL_0569;
							case 65:
								goto IL_0575;
							case 66:
								goto IL_0593;
							case 67:
								goto IL_05a9;
							case 68:
								goto IL_05cf;
							case 70:
								goto IL_05d7;
							case 64:
							case 69:
							case 71:
								goto IL_05e5;
							case 72:
								goto IL_05fa;
							case 73:
								goto IL_061f;
							case 75:
								goto IL_062e;
							case 76:
								goto IL_064c;
							case 77:
								goto IL_0662;
							case 78:
								goto IL_0682;
							case 80:
								goto IL_068a;
							case 74:
							case 79:
							case 81:
								goto IL_0698;
							case 82:
								goto IL_06ad;
							case 83:
								goto IL_06d4;
							case 85:
								goto IL_06e3;
							case 86:
								goto IL_0701;
							case 87:
								goto IL_070d;
							case 88:
								goto IL_072d;
							case 90:
								goto IL_0735;
							case 84:
							case 89:
							case 91:
								goto IL_0743;
							case 92:
								goto IL_0746;
							case 93:
								goto IL_0764;
							case 94:
								goto IL_076b;
							case 95:
								goto IL_0771;
							case 97:
								goto IL_0786;
							case 98:
								goto IL_0790;
							case 99:
								goto IL_079c;
							case 100:
								goto IL_07a8;
							case 101:
								goto IL_07b4;
							case 102:
								goto IL_07c0;
							case 103:
								goto IL_07cc;
							case 104:
								goto IL_07d8;
							case 105:
								goto IL_0824;
							case 106:
								goto IL_082a;
							case 107:
								goto IL_0830;
							case 108:
								goto IL_0846;
							case 110:
								goto IL_0895;
							case 111:
								goto IL_08c4;
							case 113:
								goto IL_08ec;
							case 114:
								goto IL_0918;
							case 116:
								goto IL_0928;
							case 117:
								goto IL_0949;
							case 119:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 96:
							case 109:
							case 112:
							case 115:
							case 118:
							case 120:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_079c:
						num2 = 99;
						list.Add(item);
						goto IL_07a8;
						IL_000a:
						num2 = 2;
						num5 = 0;
						goto IL_000f;
						IL_000f:
						num2 = 3;
						a = strFunction.Length + 2;
						goto IL_001e;
						IL_001e:
						num2 = 4;
						Evaluate.m_A = a;
						goto IL_0062;
						IL_0062:
						num2 = 6;
						if (Evaluate.m_A < strFormula.Length)
						{
							goto IL_0029;
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
						goto IL_0087;
						IL_07a8:
						num2 = 100;
						list.Add(item2);
						goto IL_07b4;
						IL_07b4:
						num2 = 101;
						list.Add(text);
						goto IL_07c0;
						IL_07c0:
						num2 = 102;
						list.Add(item3);
						goto IL_07cc;
						IL_0087:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_008e;
						IL_008e:
						num2 = 12;
						application = MH.A.Application;
						goto IL_00a1;
						IL_00a1:
						num2 = 13;
						workbook = application.ActiveWorkbook;
						goto IL_00ad;
						IL_00ad:
						num2 = 14;
						array2 = new string[2];
						goto IL_00b8;
						IL_00b8:
						num2 = 15;
						array2[0] = "";
						goto IL_00c4;
						IL_00c4:
						num2 = 16;
						name = workbook.Names.Item(RuntimeHelpers.GetObjectValue(array[1]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_00fa;
						IL_00fa:
						num2 = 17;
						if (Information.Err().Number == 0)
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
							goto IL_0117;
						}
						goto IL_014a;
						IL_07cc:
						num2 = 103;
						list.Add(item4);
						goto IL_07d8;
						IL_0117:
						num2 = 18;
						array[1] = Strings.Right(Conversions.ToString(name.RefersTo), Strings.Len(RuntimeHelpers.GetObjectValue(name.RefersTo)) - 1);
						goto IL_0158;
						IL_014a:
						num2 = 20;
						Information.Err().Number = 0;
						goto IL_0158;
						IL_0158:
						num2 = 21;
						if (Strings.InStr(1, Conversions.ToString(array[1]), VH.A(7827)) > 0)
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
							goto IL_0189;
						}
						goto IL_0216;
						IL_07d8:
						num2 = 104;
						objectValue = RuntimeHelpers.GetObjectValue(application.Evaluate(VH.A(158912) + Strings.Join(list.ToArray(), VH.A(2378)) + VH.A(39904)));
						goto IL_0824;
						IL_0189:
						num2 = 22;
						array3 = Strings.Split(Conversions.ToString(array[1]), VH.A(7827));
						goto IL_01ac;
						IL_01ac:
						num2 = 23;
						if (!Operators.ConditionalCompareObjectEqual(array3[0], NewLateBinding.LateGet(workbook.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), TextCompare: false))
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
							goto IL_01eb;
						}
						goto IL_0209;
						IL_0824:
						num2 = 105;
						application = null;
						goto IL_082a;
						IL_01eb:
						num2 = 25;
						array2[0] = array3[0] + VH.A(7827);
						goto IL_0209;
						IL_0209:
						num2 = 26;
						array2[1] = array3[1];
						goto IL_0228;
						IL_0216:
						num2 = 28;
						array2[1] = Conversions.ToString(array[1]);
						goto IL_0228;
						IL_0228:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_022f;
						IL_022f:
						num2 = 30;
						array4 = new string[2];
						goto IL_023a;
						IL_023a:
						num2 = 31;
						array4[0] = "";
						goto IL_0246;
						IL_0246:
						num2 = 32;
						name = workbook.Names.Item(RuntimeHelpers.GetObjectValue(array[2]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_027c;
						IL_027c:
						num2 = 33;
						if (Information.Err().Number == 0)
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
							goto IL_0297;
						}
						goto IL_02d0;
						IL_082a:
						num2 = 106;
						list = null;
						goto IL_0830;
						IL_0297:
						num2 = 34;
						array[2] = Strings.Right(Conversions.ToString(name.RefersTo), Strings.Len(RuntimeHelpers.GetObjectValue(name.RefersTo)) - 1);
						goto IL_02e0;
						IL_02d0:
						num2 = 36;
						Information.Err().Number = 0;
						goto IL_02e0;
						IL_02e0:
						num2 = 37;
						if (Strings.InStr(1, Conversions.ToString(array[2]), VH.A(7827)) > 0)
						{
							goto IL_0309;
						}
						goto IL_0394;
						IL_0309:
						num2 = 38;
						array3 = Strings.Split(Conversions.ToString(array[2]), VH.A(7827));
						goto IL_032c;
						IL_032c:
						num2 = 39;
						if (!Operators.ConditionalCompareObjectEqual(array3[0], NewLateBinding.LateGet(workbook.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), TextCompare: false))
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
							goto IL_036b;
						}
						goto IL_0387;
						IL_0830:
						num2 = 107;
						if (objectValue is Range)
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
							goto IL_0846;
						}
						goto IL_0895;
						IL_036b:
						num2 = 41;
						array4[0] = array3[0] + VH.A(7827);
						goto IL_0387;
						IL_0387:
						num2 = 42;
						array4[1] = array3[1];
						goto IL_03a4;
						IL_0394:
						num2 = 44;
						array4[1] = Conversions.ToString(array[2]);
						goto IL_03a4;
						IL_03a4:
						num2 = 45;
						workbook = null;
						goto IL_03aa;
						IL_03aa:
						num2 = 46;
						text = "";
						goto IL_03b4;
						IL_03b4:
						num2 = 47;
						item3 = "";
						goto IL_03be;
						IL_03be:
						num2 = 48;
						item4 = "";
						goto IL_03c8;
						IL_03c8:
						num2 = 49;
						application2 = application;
						goto IL_03cf;
						IL_03cf:
						num2 = 50;
						obj = RuntimeHelpers.GetObjectValue(application2.Evaluate(RuntimeHelpers.GetObjectValue(array[0])));
						goto IL_03ed;
						IL_03ed:
						num2 = 51;
						if (obj is Range)
						{
							goto IL_03f9;
						}
						goto IL_0443;
						IL_03f9:
						num2 = 52;
						item5 = ((Range)obj).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_043b;
						IL_043b:
						num2 = 53;
						obj = null;
						goto IL_049d;
						IL_0443:
						num2 = 55;
						if (obj is string)
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
							goto IL_0459;
						}
						goto IL_048f;
						IL_0928:
						num2 = 116;
						if (!Operators.ConditionalCompareObjectEqual(objectValue, -2146826273, TextCompare: false))
						{
							break;
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
						goto IL_0949;
						IL_0459:
						num2 = 56;
						item5 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(39830), obj), VH.A(39830)));
						goto IL_049d;
						IL_048f:
						num2 = 58;
						item5 = obj.ToString();
						goto IL_049d;
						IL_049d:
						num2 = 59;
						item = ((_Application)application2).get_Range((object)(array2[0] + array2[1]), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_04e4;
						IL_04e4:
						num2 = 60;
						item2 = ((_Application)application2).get_Range((object)(array4[0] + array4[1]), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_052f;
						IL_052f:
						num2 = 61;
						if (num5 > 3)
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
							goto IL_0544;
						}
						goto IL_0743;
						IL_0846:
						num2 = 108;
						result = array4[0] + ((Range)objectValue).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto end_IL_0000_3;
						IL_0544:
						num2 = 62;
						if (Operators.CompareString(array[3].ToString(), "", TextCompare: false) == 0)
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
							goto IL_0569;
						}
						goto IL_0575;
						IL_0895:
						num2 = 110;
						if (Operators.ConditionalCompareObjectEqual(objectValue, -2146826259, TextCompare: false))
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
							if (text.Length > 0)
							{
								goto IL_08c4;
							}
						}
						goto IL_08ec;
						IL_0569:
						num2 = 63;
						text = "";
						goto IL_05e5;
						IL_0575:
						num2 = 65;
						obj2 = RuntimeHelpers.GetObjectValue(application2.Evaluate(RuntimeHelpers.GetObjectValue(array[3])));
						goto IL_0593;
						IL_0593:
						num2 = 66;
						if (obj2 is Range)
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
							goto IL_05a9;
						}
						goto IL_05d7;
						IL_0059:
						num2 = 9;
						num5++;
						goto IL_0062;
						IL_05a9:
						num2 = 67;
						text = Conversions.ToString(((Range)obj2).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
						goto IL_05cf;
						IL_05cf:
						num2 = 68;
						obj2 = null;
						goto IL_05e5;
						IL_05d7:
						num2 = 70;
						text = obj2.ToString();
						goto IL_05e5;
						IL_05e5:
						num2 = 71;
						if (num5 > 4)
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
							goto IL_05fa;
						}
						goto IL_0743;
						IL_0949:
						num2 = 117;
						result = Conversions.ToString(objectValue);
						goto end_IL_0000_3;
						IL_05fa:
						num2 = 72;
						if (Operators.CompareString(array[4].ToString(), "", TextCompare: false) == 0)
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
							goto IL_061f;
						}
						goto IL_062e;
						IL_08c4:
						num2 = 111;
						result = VH.A(39830) + text + VH.A(39830);
						goto end_IL_0000_3;
						IL_061f:
						num2 = 73;
						item3 = Conversions.ToString(0);
						goto IL_0698;
						IL_062e:
						num2 = 75;
						obj3 = RuntimeHelpers.GetObjectValue(application2.Evaluate(RuntimeHelpers.GetObjectValue(array[4])));
						goto IL_064c;
						IL_064c:
						num2 = 76;
						if (obj3 is Range)
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
							goto IL_0662;
						}
						goto IL_068a;
						IL_08ec:
						num2 = 113;
						if (Operators.ConditionalCompareObjectEqual(objectValue, -2146826246, TextCompare: false))
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
							if (text.Length == 0)
							{
								goto IL_0918;
							}
						}
						goto IL_0928;
						IL_0662:
						num2 = 77;
						item3 = Conversions.ToString(((Range)obj3).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
						goto IL_0682;
						IL_0682:
						num2 = 78;
						obj3 = null;
						goto IL_0698;
						IL_068a:
						num2 = 80;
						item3 = obj3.ToString();
						goto IL_0698;
						IL_0698:
						num2 = 81;
						if (num5 > 5)
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
							goto IL_06ad;
						}
						goto IL_0743;
						IL_0029:
						num2 = 7;
						array = (object[])Utils.CopyArray(array, new object[num5 + 1]);
						goto IL_0044;
						IL_06ad:
						num2 = 82;
						if (Operators.CompareString(array[5].ToString(), "", TextCompare: false) == 0)
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
							goto IL_06d4;
						}
						goto IL_06e3;
						IL_0044:
						num2 = 8;
						array[num5] = C(strFormula, (int)Evaluate.m_A);
						goto IL_0059;
						IL_06d4:
						num2 = 83;
						item4 = Conversions.ToString(1);
						goto IL_0743;
						IL_06e3:
						num2 = 85;
						obj4 = RuntimeHelpers.GetObjectValue(application2.Evaluate(RuntimeHelpers.GetObjectValue(array[5])));
						goto IL_0701;
						IL_0701:
						num2 = 86;
						if (obj4 is Range)
						{
							goto IL_070d;
						}
						goto IL_0735;
						IL_070d:
						num2 = 87;
						item4 = Conversions.ToString(((Range)obj4).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
						goto IL_072d;
						IL_072d:
						num2 = 88;
						obj4 = null;
						goto IL_0743;
						IL_0735:
						num2 = 90;
						item4 = obj4.ToString();
						goto IL_0743;
						IL_0743:
						application2 = null;
						goto IL_0746;
						IL_0746:
						num2 = 92;
						if (Information.Err().Number > 0)
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
							goto IL_0764;
						}
						goto IL_0786;
						IL_0918:
						num2 = 114;
						result = VH.A(158903);
						goto end_IL_0000_3;
						IL_0764:
						ProjectData.ClearProjectError();
						num3 = 0;
						goto IL_076b;
						IL_076b:
						num2 = 94;
						application = null;
						goto IL_0771;
						IL_0771:
						num2 = 95;
						result = VH.A(158903);
						goto end_IL_0000_3;
						IL_0786:
						num2 = 97;
						list = new List<string>();
						goto IL_0790;
						IL_0790:
						num2 = 98;
						list.Add(item5);
						goto IL_079c;
						end_IL_0000_2:
						break;
					}
					num2 = 119;
					result = Conversions.ToString(objectValue);
					break;
				}
				end_IL_0000:;
			}
			catch (object obj5) when (obj5 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj5);
				try0000_dispatch = 2902;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void VLookup()
	{
		if (A())
		{
			A(VH.A(2015), EvaluateLookup);
		}
	}

	public static void HLookup()
	{
		if (!A())
		{
			return;
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
			A(VH.A(2030), EvaluateLookup);
			return;
		}
	}

	public static string EvaluateLookup(string strFormula, string strFunction)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		object[] array = default(object[]);
		int num = default(int);
		int num3 = default(int);
		long num5 = default(long);
		long a = default(long);
		string result = default(string);
		string[] array2 = default(string[]);
		object objectValue = default(object);
		Range b = default(Range);
		int c = default(int);
		bool d = default(bool);
		Name name = default(Name);
		string[] array3 = default(string[]);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						array = new object[1];
						goto IL_000a;
					case 1195:
						{
							num = num2;
							switch (num3)
							{
							case 1:
								break;
							default:
								goto end_IL_0000;
							}
							int num4 = unchecked(num + 1);
							num = 0;
							switch (num4)
							{
							case 1:
								break;
							case 2:
								goto IL_000a;
							case 3:
								goto IL_0010;
							case 4:
								goto IL_001d;
							case 7:
								goto IL_0028;
							case 8:
								goto IL_0044;
							case 9:
								goto IL_005a;
							case 5:
							case 6:
							case 10:
								goto IL_0064;
							case 11:
								goto IL_0089;
							case 12:
								goto IL_0090;
							case 13:
								goto IL_009b;
							case 14:
								goto IL_00a7;
							case 15:
								goto IL_00e8;
							case 16:
								goto IL_00fb;
							case 18:
								goto IL_0130;
							case 17:
							case 19:
								goto IL_0140;
							case 20:
								goto IL_016f;
							case 21:
								goto IL_0192;
							case 23:
								goto IL_01dc;
							case 22:
							case 24:
								goto IL_01fa;
							case 26:
								goto IL_0207;
							case 25:
							case 27:
								goto IL_0219;
							case 28:
								goto IL_022a;
							case 29:
								goto IL_0246;
							case 30:
								goto IL_026f;
							case 31:
								goto IL_028d;
							case 33:
								goto IL_02c8;
							case 32:
							case 34:
								goto IL_02e6;
							case 35:
								goto IL_02ef;
							case 37:
								goto IL_030f;
							case 36:
							case 38:
								goto IL_0315;
							case 39:
								goto IL_0318;
							case 40:
								goto IL_032c;
							case 41:
								goto IL_0333;
							case 43:
								goto IL_0348;
							case 45:
								goto IL_0388;
							case 47:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 42:
							case 44:
							case 46:
							case 48:
							case 49:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_0348:
						num2 = 43;
						if (Operators.CompareString(strFunction, VH.A(2015), TextCompare: false) != 0)
						{
							if (Operators.CompareString(strFunction, VH.A(2030), TextCompare: false) == 0)
							{
								break;
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0000_3;
								}
							}
						}
						goto IL_0388;
						IL_000a:
						num2 = 2;
						num5 = 0L;
						goto IL_0010;
						IL_0010:
						num2 = 3;
						a = Strings.Len(strFunction) + 2;
						goto IL_001d;
						IL_001d:
						num2 = 4;
						Evaluate.m_A = a;
						goto IL_0064;
						IL_0064:
						num2 = 6;
						if (Evaluate.m_A < Strings.Len(strFormula))
						{
							goto IL_0028;
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
						goto IL_0089;
						IL_0044:
						num2 = 8;
						array[(int)num5] = C(strFormula, (int)Evaluate.m_A);
						goto IL_005a;
						IL_005a:
						num2 = 9;
						num5++;
						goto IL_0064;
						IL_0388:
						num2 = 45;
						result = array2[0] + A(RuntimeHelpers.GetObjectValue(objectValue), b, c, d);
						goto end_IL_0000_3;
						IL_0089:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0090;
						IL_0090:
						num2 = 12;
						array2 = new string[2];
						goto IL_009b;
						IL_009b:
						num2 = 13;
						array2[0] = "";
						goto IL_00a7;
						IL_00a7:
						num2 = 14;
						name = MH.A.Application.ActiveWorkbook.Names.Item(RuntimeHelpers.GetObjectValue(array[1]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_00e8;
						IL_00e8:
						num2 = 15;
						if (Information.Err().Number == 0)
						{
							goto IL_00fb;
						}
						goto IL_0130;
						IL_00fb:
						num2 = 16;
						array[1] = Strings.Right(Conversions.ToString(name.RefersTo), Strings.Len(RuntimeHelpers.GetObjectValue(name.RefersTo)) - 1);
						goto IL_0140;
						IL_0130:
						num2 = 18;
						Information.Err().Number = 0;
						goto IL_0140;
						IL_0140:
						num2 = 19;
						if (Strings.InStr(1, Conversions.ToString(array[1]), VH.A(7827)) > 0)
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
							goto IL_016f;
						}
						goto IL_0207;
						IL_0028:
						num2 = 7;
						array = (object[])Utils.CopyArray(array, new object[(int)num5 + 1]);
						goto IL_0044;
						IL_016f:
						num2 = 20;
						array3 = Strings.Split(Conversions.ToString(array[1]), VH.A(7827));
						goto IL_0192;
						IL_0192:
						num2 = 21;
						if (!Operators.ConditionalCompareObjectEqual(array3[0], NewLateBinding.LateGet(MH.A.Application.ActiveWorkbook.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), TextCompare: false))
						{
							goto IL_01dc;
						}
						goto IL_01fa;
						IL_01dc:
						num2 = 23;
						array2[0] = array3[0] + VH.A(7827);
						goto IL_01fa;
						IL_01fa:
						num2 = 24;
						array2[1] = array3[1];
						goto IL_0219;
						IL_0207:
						num2 = 26;
						array2[1] = Conversions.ToString(array[1]);
						goto IL_0219;
						IL_0219:
						num2 = 27;
						application = MH.A.Application;
						goto IL_022a;
						IL_022a:
						num2 = 28;
						objectValue = RuntimeHelpers.GetObjectValue(application.Evaluate(RuntimeHelpers.GetObjectValue(array[0])));
						goto IL_0246;
						IL_0246:
						num2 = 29;
						b = ((_Application)application).get_Range((object)(array2[0] + array2[1]), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_026f;
						IL_026f:
						num2 = 30;
						if (application.Evaluate(RuntimeHelpers.GetObjectValue(array[2])) is Range)
						{
							goto IL_028d;
						}
						goto IL_02c8;
						IL_028d:
						num2 = 31;
						c = Conversions.ToInteger(NewLateBinding.LateGet(application.Evaluate(RuntimeHelpers.GetObjectValue(array[2])), null, VH.A(41636), new object[0], null, null, null));
						goto IL_02e6;
						IL_02c8:
						num2 = 33;
						c = Conversions.ToInteger(application.Evaluate(RuntimeHelpers.GetObjectValue(array[2])));
						goto IL_02e6;
						IL_02e6:
						num2 = 34;
						if (num5 == 4)
						{
							goto IL_02ef;
						}
						goto IL_030f;
						IL_02ef:
						num2 = 35;
						d = Conversions.ToBoolean(application.Evaluate(RuntimeHelpers.GetObjectValue(array[3])));
						goto IL_0315;
						IL_030f:
						num2 = 37;
						d = true;
						goto IL_0315;
						IL_0315:
						application = null;
						goto IL_0318;
						IL_0318:
						num2 = 39;
						if (Information.Err().Number > 0)
						{
							goto IL_032c;
						}
						goto IL_0348;
						IL_032c:
						ProjectData.ClearProjectError();
						num3 = 0;
						goto IL_0333;
						IL_0333:
						num2 = 41;
						result = VH.A(44119);
						goto end_IL_0000_3;
						end_IL_0000_2:
						break;
					}
					num2 = 47;
					result = array2[0] + B(RuntimeHelpers.GetObjectValue(objectValue), b, c, d);
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1195;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static string A(object A, Range B, int C, bool D)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string text = default(string);
		long num5 = default(long);
		string result = default(string);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 330:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0017;
						case 4:
							goto IL_005d;
						case 5:
							goto IL_0085;
						case 6:
							goto IL_010c;
						case 7:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 8:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0085:
					num2 = 5;
					text = Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(B[num5, 1], null, VH.A(60565), new object[2]
					{
						0,
						checked(C - 1)
					}, null, null, null), null, VH.A(5814), new object[0], null, null, null)), VH.A(41262), "");
					goto IL_010c;
					IL_0007:
					num2 = 2;
					text = VH.A(44119);
					goto IL_0017;
					IL_0017:
					num2 = 3;
					num5 = checked((long)Math.Round(B.Application.WorksheetFunction.Match(RuntimeHelpers.GetObjectValue(A), B.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)1), D)));
					goto IL_005d;
					IL_005d:
					num2 = 4;
					if (Information.Err().Number == 0)
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
						goto IL_0085;
					}
					goto IL_010c;
					IL_010c:
					ProjectData.ClearProjectError();
					num3 = 0;
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 7;
				result = text;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 330;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static string B(object A, Range B, int C, bool D)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string text = default(string);
		long num5 = default(long);
		string result = default(string);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 330:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0017;
						case 4:
							goto IL_005f;
						case 5:
							goto IL_0087;
						case 6:
							goto IL_010c;
						case 7:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 8:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0087:
					num2 = 5;
					text = Strings.Replace(Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(B[1, num5], null, VH.A(60565), new object[2]
					{
						checked(C - 1),
						0
					}, null, null, null), null, VH.A(5814), new object[0], null, null, null)), VH.A(41262), "");
					goto IL_010c;
					IL_0007:
					num2 = 2;
					text = VH.A(44119);
					goto IL_0017;
					IL_0017:
					num2 = 3;
					num5 = checked((long)Math.Round(B.Application.WorksheetFunction.Match(RuntimeHelpers.GetObjectValue(A), B.get_Resize((object)1, RuntimeHelpers.GetObjectValue(Missing.Value)), D)));
					goto IL_005f;
					IL_005f:
					num2 = 4;
					if (Information.Err().Number == 0)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_0087;
					}
					goto IL_010c;
					IL_010c:
					ProjectData.ClearProjectError();
					num3 = 0;
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 7;
				result = text;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 330;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void IndexMatch()
	{
		if (!Evaluate.A())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		int E = default(int);
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
			Range range = null;
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.Selection is Range)
			{
				while (true)
				{
					Worksheet worksheet;
					List<Range> D;
					Range range2;
					switch (3)
					{
					case 0:
						break;
					default:
						{
							range2 = JH.A((Range)application.Selection);
							if (Conversions.ToBoolean(application.ActiveCell.HasArray))
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										Forms.WarningMessage(VH.A(158117));
										application = null;
										return;
									}
								}
							}
							range = B(range2);
							if (range != null)
							{
								D = new List<Range>();
								worksheet = range2.Worksheet;
								application.ScreenUpdating = false;
								application.EnableEvents = false;
								try
								{
									bool flag = JH.A(range2);
									enumerator = range.GetEnumerator();
									try
									{
										while (enumerator.MoveNext())
										{
											Range range3 = (Range)enumerator.Current;
											if (Conversions.ToBoolean(Operators.NotObject(range3.HasArray)))
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
												string text = range3.Formula.ToString();
												if (B(text, VH.A(4576)) > 0)
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
													if (B(text, VH.A(2045)) > 0)
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
														object objectValue = RuntimeHelpers.GetObjectValue(range3.Value2);
														string A = text;
														Evaluate.A(ref A, range3, worksheet, ref D);
														Evaluate.A(range3, RuntimeHelpers.GetObjectValue(objectValue), A, text, ref E, ref D);
													}
												}
											}
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_01ae;
											}
											continue;
											end_IL_01ae:
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
									if (flag)
									{
										JH.A(range2, VH.A(158931));
									}
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									clsReporting.LogException(ex2);
									ProjectData.ClearProjectError();
								}
								application.ScreenUpdating = true;
								application.EnableEvents = true;
								int count = D.Count;
								if (E <= 0)
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
									if (count <= 0)
									{
										Forms.InfoMessage(VH.A(159430));
										goto IL_0350;
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
								}
								if (count == 0)
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
									Forms.SuccessMessage(VH.A(158966) + E + VH.A(159065));
								}
								else
								{
									D = D.Distinct().ToList();
									if (E > 0)
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
										if (Evaluate.A(VH.A(158966) + E + VH.A(159092) + count + VH.A(159131)))
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
											Miscellaneous.GoToDialog(D);
										}
									}
									else if (Evaluate.A(VH.A(159234) + E + VH.A(159325)))
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
										Miscellaneous.GoToDialog(D);
									}
								}
								goto IL_0350;
							}
							Forms.InfoMessage(VH.A(158821));
							goto IL_036a;
						}
						IL_0350:
						range = null;
						worksheet = null;
						D = null;
						goto IL_036a;
						IL_036a:
						Evaluate.A(VH.A(159549));
						application = null;
						range2 = null;
						return;
					}
				}
			}
			Forms.WarningMessage(VH.A(158034));
			application = null;
			return;
		}
	}

	private static MatchCollection A(string A)
	{
		return Regex.Matches(A, VH.A(159574));
	}

	private static void A(ref string A, Range B, Worksheet C, ref List<Range> D)
	{
		int num = 0;
		string text = A;
		MatchCollection matchCollection = Evaluate.A(A);
		checked
		{
			while (matchCollection.Count > 0)
			{
				int num2 = matchCollection.Count - 1;
				if (num2 >= 0)
				{
					object obj = null;
					Group obj2 = matchCollection[num2].Groups[1];
					obj = RuntimeHelpers.GetObjectValue(C.Evaluate(obj2.ToString()));
					if (!(obj is Range))
					{
						D.Add(B);
						return;
					}
					Range range = (Range)obj;
					string newValue;
					if (range.Worksheet == C)
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
						newValue = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					else
					{
						newValue = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					range = null;
					A = A.Replace(A.Substring(obj2.Index, obj2.Length), newValue);
					matchCollection = Evaluate.A(A);
				}
				num++;
				if (num > 10)
				{
					A = text;
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
	}

	public static void SumIf()
	{
		if (!A())
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
			A(VH.A(4533), EvaluateSumIf);
			return;
		}
	}

	public static string EvaluateSumIf(string strFormula, string strFunction)
	{
		int num = 0;
		List<string> list = new List<string>();
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			try
			{
				object[] array = new object[1];
				string text = strFormula;
				Evaluate.m_A = strFunction.Length + 2;
				while (Evaluate.m_A < strFormula.Length)
				{
					array = (object[])Utils.CopyArray(array, new object[num + 1]);
					array[num] = C(strFormula, (int)Evaluate.m_A);
					num++;
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
					Range range = (Range)application.Evaluate(RuntimeHelpers.GetObjectValue(array[0]));
					object obj = ((!(application.Evaluate(RuntimeHelpers.GetObjectValue(array[1])) is Range)) ? Conversions.ToString(array[1]).Trim('"') : RuntimeHelpers.GetObjectValue(application.Evaluate(RuntimeHelpers.GetObjectValue(array[1]))));
					if (num == 2)
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
						try
						{
							enumerator = range.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Range range2 = (Range)enumerator.Current;
								double? num2 = A(RuntimeHelpers.GetObjectValue(range2.Value2));
								if (num2.HasValue)
								{
									double num3 = application.WorksheetFunction.SumIf(range2, RuntimeHelpers.GetObjectValue(obj), RuntimeHelpers.GetObjectValue(Missing.Value));
									if ((num2.HasValue ? new bool?(num3 == num2.GetValueOrDefault()) : ((bool?)null)) == true)
									{
										list.Add(range2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
									}
								}
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_01db;
								}
								continue;
								end_IL_01db:
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
					}
					else
					{
						string name = Conversions.ToString(array[2]);
						Range range3 = (Range)application.Evaluate(name);
						long num4 = Conversions.ToLong(range.Cells.CountLarge);
						for (long num5 = 1L; num5 <= num4; num5++)
						{
							double? num2 = A(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range3.Cells[num5, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[0], null, null, null)));
							if (!num2.HasValue)
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
							double num3 = application.WorksheetFunction.SumIf((Range)range.Cells[num5, RuntimeHelpers.GetObjectValue(Missing.Value)], RuntimeHelpers.GetObjectValue(obj), RuntimeHelpers.GetObjectValue(range3.Cells[num5, RuntimeHelpers.GetObjectValue(Missing.Value)]));
							bool? obj2;
							if (!num2.HasValue)
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
								obj2 = null;
							}
							else
							{
								obj2 = num3 == num2.GetValueOrDefault();
							}
							bool? flag = obj2;
							if (flag != true)
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
							list.Add(Conversions.ToString(NewLateBinding.LateGet(range3.Cells[num5, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)));
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
					}
					if (list.Any())
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								return Strings.Join(list.ToArray(), VH.A(54459));
							}
						}
					}
					return VH.A(64024);
				}
			}
			finally
			{
				list = null;
				application = null;
				Range range = null;
				Range range3 = null;
			}
		}
	}

	public static void SumIfs()
	{
		if (A())
		{
			A(VH.A(49109), EvaluateSumIfs);
		}
	}

	public static string EvaluateSumIfs(string strFormula, string strFunction)
	{
		int num = 0;
		List<string> list = new List<string>();
		new List<Range>();
		new List<string>();
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			try
			{
				object[] array = new object[1];
				string text = strFormula;
				Evaluate.m_A = strFunction.Length + 2;
				while (Evaluate.m_A < strFormula.Length)
				{
					array = (object[])Utils.CopyArray(array, new object[num + 1]);
					array[num] = C(strFormula, (int)Evaluate.m_A);
					num++;
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
					string name = Conversions.ToString(array[0]);
					Range range = (Range)application.Evaluate(name);
					object[,] array2 = new object[1, 1];
					array2 = new object[14, 2];
					int num2 = (int)Math.Round((double)(num - 1) / 2.0);
					try
					{
						int num3 = Information.UBound(array2) - 1;
						for (int i = 0; i <= num3; i++)
						{
							if (i + 1 > num2)
							{
								array2[i, 0] = Missing.Value;
								array2[i, 1] = Missing.Value;
								continue;
							}
							int num4 = 1 + i * 2;
							array2[i, 0] = RuntimeHelpers.GetObjectValue(application.Evaluate(RuntimeHelpers.GetObjectValue(array[num4])));
							if (application.Evaluate(RuntimeHelpers.GetObjectValue(array[num4 + 1])) is Range)
							{
								array2[i, 1] = RuntimeHelpers.GetObjectValue(application.Evaluate(RuntimeHelpers.GetObjectValue(array[num4 + 1])));
								continue;
							}
							string text2 = Conversions.ToString(array[num4 + 1]);
							array2[i, 1] = text2.Trim('"');
						}
						long num5 = Conversions.ToLong(range.Cells.CountLarge);
						for (long num6 = 1L; num6 <= num5; num6++)
						{
							double? num7 = A(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range.Cells[num6, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[0], null, null, null)));
							if (!num7.HasValue)
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
							object[] array3 = new object[Information.UBound(array2) + 1];
							int num8 = Information.UBound(array2);
							for (int num4 = 0; num4 <= num8; num4++)
							{
								array3[num4] = RuntimeHelpers.GetObjectValue(A(RuntimeHelpers.GetObjectValue(array2[num4, 0]), (int)num6));
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
							double num9 = application.WorksheetFunction.SumIfs((Range)range.Cells[num6, RuntimeHelpers.GetObjectValue(Missing.Value)], (Range)array3[0], RuntimeHelpers.GetObjectValue(array2[0, 1]), RuntimeHelpers.GetObjectValue(array3[1]), RuntimeHelpers.GetObjectValue(array2[1, 1]), RuntimeHelpers.GetObjectValue(array3[2]), RuntimeHelpers.GetObjectValue(array2[2, 1]), RuntimeHelpers.GetObjectValue(array3[3]), RuntimeHelpers.GetObjectValue(array2[3, 1]), RuntimeHelpers.GetObjectValue(array3[4]), RuntimeHelpers.GetObjectValue(array2[4, 1]), RuntimeHelpers.GetObjectValue(array3[5]), RuntimeHelpers.GetObjectValue(array2[5, 1]), RuntimeHelpers.GetObjectValue(array3[6]), RuntimeHelpers.GetObjectValue(array2[6, 1]), RuntimeHelpers.GetObjectValue(array3[7]), RuntimeHelpers.GetObjectValue(array2[7, 1]), RuntimeHelpers.GetObjectValue(array3[8]), RuntimeHelpers.GetObjectValue(array2[8, 1]), RuntimeHelpers.GetObjectValue(array3[9]), RuntimeHelpers.GetObjectValue(array2[9, 1]), RuntimeHelpers.GetObjectValue(array3[10]), RuntimeHelpers.GetObjectValue(array2[10, 1]), RuntimeHelpers.GetObjectValue(array3[11]), RuntimeHelpers.GetObjectValue(array2[11, 1]), RuntimeHelpers.GetObjectValue(array3[12]), RuntimeHelpers.GetObjectValue(array2[12, 1]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							if ((num7.HasValue ? new bool?(num9 == num7.GetValueOrDefault()) : ((bool?)null)) == true)
							{
								list.Add(Conversions.ToString(NewLateBinding.LateGet(range.Cells[num6, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)));
							}
							array3 = null;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_04ff;
							}
							continue;
							end_IL_04ff:
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Interaction.MsgBox(ex2.Message);
						throw;
					}
					if (list.Any())
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								return Strings.Join(list.ToArray(), VH.A(54459));
							}
						}
					}
					return VH.A(64024);
				}
			}
			finally
			{
				list = null;
				application = null;
				Range range = null;
			}
		}
	}

	private static object A(object A, int B)
	{
		if (A is Range)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return ((Range)A).Cells[B, RuntimeHelpers.GetObjectValue(Missing.Value)];
				}
			}
		}
		return A;
	}

	private static void A(string A, Func<string, string, string> B)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		checked
		{
			double num = default(double);
			IEnumerator enumerator = default(IEnumerator);
			int num3 = default(int);
			bool flag2 = default(bool);
			if (application.Selection is Range)
			{
				while (true)
				{
					Range a;
					Range range;
					List<Range> list;
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
							a = (Range)application.Selection;
							if (Conversions.ToBoolean(application.ActiveCell.HasArray))
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										Forms.WarningMessage(VH.A(158117));
										application = null;
										return;
									}
								}
							}
							range = Evaluate.B(a);
							if (range != null)
							{
								list = new List<Range>();
								try
								{
									num = application.WorksheetFunction.Sum(range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
								application.ScreenUpdating = false;
								application.EnableEvents = false;
								try
								{
									bool flag = JH.A(a);
									try
									{
										enumerator = range.GetEnumerator();
										while (enumerator.MoveNext())
										{
											Range range2 = (Range)enumerator.Current;
											if (Conversions.ToBoolean(Operators.NotObject(range2.HasArray)))
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
												string text = Conversions.ToString(range2.Formula);
												if (Evaluate.B(text, A) > 0)
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
													RuntimeHelpers.GetObjectValue(range2.Value2);
													text = Strings.Right(text, Strings.Len(text) - 1);
													int num2 = Evaluate.B(text, A);
													while (true)
													{
														if (num2 <= 0)
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
															if (Operators.CompareString(A, VH.A(158202), TextCompare: false) != 0)
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
																if (Strings.InStr(1, text, VH.A(44119)) > 0)
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
																	list.Add(range2);
																	break;
																}
															}
															string text2 = VH.A(48936);
															if (Versioned.IsNumeric(text))
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
																text2 = "";
															}
															try
															{
																range2.Formula = text2 + text;
															}
															catch (Exception ex3)
															{
																ProjectData.SetProjectError(ex3);
																Exception ex4 = ex3;
																list.Add(range2);
																ProjectData.ClearProjectError();
																break;
															}
															text = Conversions.ToString(range2.Formula);
															string left = Strings.Mid(text, 2, 1);
															string left2 = Strings.Right(text, 1);
															string text3 = Strings.Replace(text, VH.A(39830), "");
															if (Operators.CompareString(left, VH.A(39830), TextCompare: false) == 0 && Operators.CompareString(left2, VH.A(39830), TextCompare: false) == 0)
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
																if (text.Length - text3.Length == 2)
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
																	try
																	{
																		range2.Formula = RuntimeHelpers.GetObjectValue(application.Evaluate(text));
																	}
																	catch (Exception ex5)
																	{
																		ProjectData.SetProjectError(ex5);
																		Exception ex6 = ex5;
																		list.Add(range2);
																		ProjectData.ClearProjectError();
																		break;
																	}
																}
															}
															num3++;
															break;
														}
														string text4 = D(text, num2);
														string replacement;
														try
														{
															replacement = B(text4, A);
														}
														catch (Exception ex7)
														{
															ProjectData.SetProjectError(ex7);
															Exception ex8 = ex7;
															list.Add(range2);
															ProjectData.ClearProjectError();
															break;
														}
														text = Strings.Replace(text, text4, replacement, 1, 1);
														num2 = Evaluate.B(text, A);
													}
												}
											}
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_049a;
											}
											continue;
											end_IL_049a:
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
													break;
												default:
													(enumerator as IDisposable).Dispose();
													goto end_IL_04af;
												}
												continue;
												end_IL_04af:
												break;
											}
										}
									}
									if (flag)
									{
										JH.A(a, VH.A(158931));
									}
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									ProjectData.ClearProjectError();
								}
								application.ScreenUpdating = true;
								application.EnableEvents = true;
								try
								{
									if (num != application.WorksheetFunction.Sum(range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) && Evaluate.A(VH.A(159655)))
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												Core.Undo();
												flag2 = true;
												goto end_IL_0672;
											}
											continue;
											end_IL_0672:
											break;
										}
									}
								}
								catch (Exception ex11)
								{
									ProjectData.SetProjectError(ex11);
									Exception ex12 = ex11;
									ProjectData.ClearProjectError();
								}
								if (!flag2)
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
									if (num3 <= 0)
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
										if (!list.Any())
										{
											Forms.InfoMessage(VH.A(159966) + A + VH.A(160031));
											goto IL_07e1;
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
									}
									if (!list.Any())
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
										Forms.SuccessMessage(VH.A(159853) + A + VH.A(159900) + num3 + VH.A(158409));
									}
									else if (Evaluate.A(VH.A(159853) + A + VH.A(159900) + num3 + VH.A(159929) + list.Count + VH.A(159131)))
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
										Miscellaneous.GoToDialog(list);
									}
								}
								goto IL_07e1;
							}
							Forms.InfoMessage(VH.A(158821));
							goto IL_07f9;
						}
						IL_07f9:
						Evaluate.A(A);
						application = null;
						a = null;
						return;
						IL_07e1:
						range = null;
						list = null;
						goto IL_07f9;
					}
				}
			}
			Forms.WarningMessage(VH.A(158034));
			application = null;
		}
	}

	private static double? A(object A)
	{
		return (!JH.A(RuntimeHelpers.GetObjectValue(A))) ? ((double?)null) : new double?(Conversions.ToDouble(A));
	}

	private static void A(List<double> A, object B, bool C, ref string D, Func<string> E)
	{
		double? A2 = Evaluate.A(RuntimeHelpers.GetObjectValue(B));
		double num;
		if (!C)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			num = -1.0;
		}
		else
		{
			num = 1.0;
		}
		double A3 = num;
		if (!A2.HasValue)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (!A.Any())
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								D = E();
								return;
							}
						}
					}
					return;
				}
			}
		}
		if (A.Any())
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
			if (A.Where([SpecialName] (double num2) => A3 * num2 < A3 * A2.Value).Any())
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
				break;
			}
		}
		D = E();
		A.Add(A2.Value);
	}

	private static int B(string A, string B)
	{
		return Strings.InStr(1, A, B);
	}

	private static long A(object A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (application.Evaluate(RuntimeHelpers.GetObjectValue(A)) is Range)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (application.WorksheetFunction.IsError(RuntimeHelpers.GetObjectValue(application.Evaluate(RuntimeHelpers.GetObjectValue(A)))))
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								throw new Exception();
							}
						}
					}
					return Conversions.ToLong(NewLateBinding.LateGet(application.Evaluate(RuntimeHelpers.GetObjectValue(A)), null, VH.A(41636), new object[0], null, null, null));
				}
			}
		}
		return Conversions.ToLong(application.Evaluate(RuntimeHelpers.GetObjectValue(A)));
	}

	private static string C(string A, int B)
	{
		string text = "";
		string right = VH.A(39848);
		string right2 = VH.A(39904);
		string right3 = VH.A(2378);
		int num = B;
		int num2 = 0;
		checked
		{
			bool flag = default(bool);
			while (true)
			{
				if (num <= A.Length)
				{
					flag = true;
					string text2 = Strings.Mid(A, num, 1);
					string left = text2;
					if (Operators.CompareString(left, right, TextCompare: false) == 0)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						num2++;
						flag = true;
					}
					else if (Operators.CompareString(left, right2, TextCompare: false) == 0)
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
						num2--;
					}
					else if (Operators.CompareString(left, right3, TextCompare: false) == 0)
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
						if (num2 == 0)
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
							flag = false;
							break;
						}
					}
					text += text2;
					num++;
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
				break;
			}
			Evaluate.m_A = Evaluate.m_A + Strings.Len(text) + 1;
			if (flag)
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
				text = Strings.Left(text, text.Length - 1);
			}
			return text;
		}
	}

	private static string D(string A, int B)
	{
		string text = "";
		string right = VH.A(39848);
		string right2 = VH.A(39904);
		int num = B;
		int num2 = 0;
		checked
		{
			while (true)
			{
				if (num <= A.Length)
				{
					string text2 = Strings.Mid(A, num, 1);
					text += text2;
					if (Operators.CompareString(text2, right, TextCompare: false) == 0)
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
						num2++;
					}
					if (Operators.CompareString(text2, right2, TextCompare: false) == 0)
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
						num2--;
						if (num2 == 0)
						{
							break;
						}
					}
					num++;
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
			return text;
		}
	}

	private static void A(Range A, object B, string C, string D, ref int E, ref List<Range> F)
	{
		if (Operators.CompareString(D, C, TextCompare: false) == 0)
		{
			return;
		}
		checked
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
				try
				{
					C = C.Replace(VH.A(3799), VH.A(54459));
					C = C.Replace(VH.A(160080), VH.A(54459));
					A.Formula = C;
					if (Operators.ConditionalCompareObjectEqual(A.Value2, B, TextCompare: false))
					{
						E++;
						return;
					}
					A.Formula = D;
					F.Add(A);
					return;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					A.Formula = D;
					F.Add(A);
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	private static Range B(Range A)
	{
		Range result = null;
		if (Operators.ConditionalCompareObjectEqual(A.Cells.CountLarge, 1, TextCompare: false))
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
			if (Conversions.ToBoolean(A.HasFormula))
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
				result = A;
			}
		}
		else
		{
			try
			{
				result = A.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private static bool A(string A)
	{
		return Forms.YesNoMessage2(A, (YesNoDefault)1) == DialogResult.Yes;
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(160085) + A);
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
