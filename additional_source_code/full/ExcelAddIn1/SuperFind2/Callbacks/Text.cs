using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Text
{
	internal static void A(WorksheetItem A, Range B)
	{
		Text.A(A, B, (Func<string, string, bool>)Text.A);
	}

	private static bool A(string A, string B)
	{
		return Operators.CompareString(A, B, TextCompare: false) == 0;
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Text.A(A, B, (Func<string, string, bool>)Text.B);
	}

	private static bool B(string A, string B)
	{
		return A.Contains(B);
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Text.A(A, B, C);
	}

	private static bool C(string A, string B)
	{
		return !A.Contains(B);
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Text.A(A, B, D);
	}

	private static bool D(string A, string B)
	{
		return A.StartsWith(B);
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Text.A(A, B, E);
	}

	private static bool E(string A, string B)
	{
		return A.EndsWith(B);
	}

	internal static void F(WorksheetItem A, Range B)
	{
		Text.A(A, B, F);
	}

	private static bool F(string A, string B)
	{
		return !A.StartsWith(B);
	}

	internal static void G(WorksheetItem A, Range B)
	{
		Text.A(A, B, G);
	}

	private static bool G(string A, string B)
	{
		return !A.EndsWith(B);
	}

	private static void A(WorksheetItem A, Range B, Func<string, string, bool> C)
	{
		Range range = null;
		string text = Props.SearchForm.Input1;
		bool matchCase = Props.SearchForm.MatchCase;
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
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
			if (!matchCase)
			{
				text = text.ToLower();
			}
			if (!Props.SearchForm.LookInValues && !Props.SearchForm.LookInFormulas)
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
				if (!Props.SearchForm.LookInHyperlinks)
				{
					goto IL_0082;
				}
			}
			range = RangeHelpers.H(B);
			goto IL_0082;
			IL_0082:
			if (Props.SearchForm.LookInValues && range != null)
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
				object objectValue;
				try
				{
					enumerator = range.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range a = (Range)enumerator.Current;
						try
						{
							objectValue = RuntimeHelpers.GetObjectValue(Text.A(a));
							string arg = (matchCase ? objectValue.ToString() : objectValue.ToString().ToLower());
							if (!C(arg, text))
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
								A.A(a);
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
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_012a;
						}
						continue;
						end_IL_012a:
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
				objectValue = null;
			}
			if (Props.SearchForm.LookInFormulas)
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
				if (range != null)
				{
					try
					{
						enumerator2 = range.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Range range2 = (Range)enumerator2.Current;
							try
							{
								object obj = NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null);
								string arg = (matchCase ? obj.ToString() : obj.ToString().ToLower());
								obj = null;
								if (!C(arg, text))
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
									A.A(range2);
									break;
								}
								continue;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0215;
							}
							continue;
							end_IL_0215:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (2)
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
			}
			if (Props.SearchForm.LookInComments)
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
				try
				{
					Range range3 = RangeHelpers.F(B);
					if (range3 != null)
					{
						{
							enumerator3 = range3.GetEnumerator();
							try
							{
								while (enumerator3.MoveNext())
								{
									Range range4 = (Range)enumerator3.Current;
									try
									{
										Comment comment = range4.Comment;
										string text2;
										if (!matchCase)
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
											text2 = comment.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).ToLower();
										}
										else
										{
											text2 = comment.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										string arg = text2;
										comment = null;
										if (C(arg, text))
										{
											A.A(range4);
										}
									}
									catch (Exception ex5)
									{
										ProjectData.SetProjectError(ex5);
										Exception ex6 = ex5;
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
										goto end_IL_0333;
									}
									continue;
									end_IL_0333:
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
						range3 = null;
					}
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					ProjectData.ClearProjectError();
				}
				try
				{
					List<Range> list = RangeHelpers.A(B);
					if (list != null)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							using (List<Range>.Enumerator enumerator4 = list.GetEnumerator())
							{
								while (enumerator4.MoveNext())
								{
									Range current = enumerator4.Current;
									try
									{
										object instance = NewLateBinding.LateGet(current, null, VH.A(103833), new object[0], null, null, null);
										object obj2;
										if (!matchCase)
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
											obj2 = NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null);
										}
										else
										{
											obj2 = NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null);
										}
										string arg = Conversions.ToString(obj2);
										if (C(arg, text))
										{
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												A.A(current);
												break;
											}
											continue;
										}
										arg = Conversions.ToString(matchCase ? NewLateBinding.LateGet(instance, null, VH.A(96399), new object[0], null, null, null) : NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(96399), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null));
										if (C(arg, text))
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												A.A(current);
												break;
											}
											continue;
										}
										int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
										for (int i = 1; i <= num; i = checked(i + 1))
										{
											object[] array;
											bool[] array2;
											object obj3 = NewLateBinding.LateGet(instance, null, VH.A(102647), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
											if (array2[0])
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
												i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
											}
											object instance2 = obj3;
											arg = Conversions.ToString(matchCase ? NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null) : NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null));
											if (C(arg, text))
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
												A.A(current);
												break;
											}
											object obj4;
											if (!matchCase)
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
												obj4 = NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(96399), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null);
											}
											else
											{
												obj4 = NewLateBinding.LateGet(instance2, null, VH.A(96399), new object[0], null, null, null);
											}
											arg = Conversions.ToString(obj4);
											if (C(arg, text))
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
												A.A(current);
												break;
											}
											instance2 = null;
										}
										instance = null;
									}
									catch (Exception ex9)
									{
										ProjectData.SetProjectError(ex9);
										Exception ex10 = ex9;
										ProjectData.ClearProjectError();
									}
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0746;
									}
									continue;
									end_IL_0746:
									break;
								}
							}
							list = null;
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
			}
			if (Props.SearchForm.LookInCharts)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						throw new NotImplementedException();
					}
				}
			}
			if (Props.SearchForm.LookInHyperlinks)
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
				if (range != null)
				{
					try
					{
						enumerator5 = range.GetEnumerator();
						while (enumerator5.MoveNext())
						{
							Range range5 = (Range)enumerator5.Current;
							foreach (Hyperlink hyperlink4 in range5.Hyperlinks)
							{
								Hyperlink hyperlink2 = hyperlink4;
								string text3;
								if (!matchCase)
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
									text3 = hyperlink2.Address.ToLower();
								}
								else
								{
									text3 = hyperlink2.Address;
								}
								string arg = text3;
								hyperlink2 = null;
								if (C(arg, text))
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										A.A(range5);
										break;
									}
									break;
								}
								Hyperlink hyperlink3 = hyperlink4;
								string text4;
								if (!matchCase)
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
									text4 = hyperlink3.TextToDisplay.ToLower();
								}
								else
								{
									text4 = hyperlink3.TextToDisplay;
								}
								arg = text4;
								hyperlink3 = null;
								if (!C(arg, text))
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
									A.A(range5);
									break;
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
								goto end_IL_08ce;
							}
							continue;
							end_IL_08ce:
							break;
						}
					}
					finally
					{
						if (enumerator5 is IDisposable)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								(enumerator5 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			range = null;
			return;
		}
	}

	internal static void H(WorksheetItem A, Range B)
	{
		Range range = null;
		string input = Props.SearchForm.Input1;
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
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
			if (!Props.SearchForm.LookInValues)
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
				if (!Props.SearchForm.LookInFormulas)
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
					if (!Props.SearchForm.LookInHyperlinks)
					{
						goto IL_007d;
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
				}
			}
			range = RangeHelpers.H(B);
			goto IL_007d;
			IL_007d:
			if (Props.SearchForm.LookInValues)
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
					object objectValue;
					try
					{
						enumerator = range.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Range a = (Range)enumerator.Current;
							objectValue = RuntimeHelpers.GetObjectValue(Text.A(a));
							try
							{
								if (!Regex.IsMatch(objectValue.ToString(), input))
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
									A.A(a);
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
					objectValue = null;
				}
			}
			if (Props.SearchForm.LookInFormulas)
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
					try
					{
						enumerator2 = range.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Range range2 = (Range)enumerator2.Current;
							try
							{
								if (Regex.IsMatch(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null).ToString(), input))
								{
									A.A(range2);
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_01cb;
							}
							continue;
							end_IL_01cb:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (3)
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
			}
			if (Props.SearchForm.LookInComments)
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
				Range range3 = RangeHelpers.F(B);
				if (range3 != null)
				{
					try
					{
						enumerator3 = range3.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							Range range4 = (Range)enumerator3.Current;
							try
							{
								if (!Regex.IsMatch(range4.Comment.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), input))
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
									A.A(range4);
									break;
								}
								continue;
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_02a8;
							}
							continue;
							end_IL_02a8:
							break;
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator3 as IDisposable).Dispose();
								break;
							}
						}
					}
					range3 = null;
				}
				try
				{
					List<Range> list = RangeHelpers.A(B);
					if (list != null)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							using (List<Range>.Enumerator enumerator4 = list.GetEnumerator())
							{
								while (enumerator4.MoveNext())
								{
									Range current = enumerator4.Current;
									try
									{
										object instance = NewLateBinding.LateGet(current, null, VH.A(103833), new object[0], null, null, null);
										if (Regex.IsMatch(Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null)), input))
										{
											while (true)
											{
												switch (7)
												{
												case 0:
													continue;
												}
												A.A(current);
												break;
											}
											continue;
										}
										if (Regex.IsMatch(Conversions.ToString(NewLateBinding.LateGet(instance, null, VH.A(96399), new object[0], null, null, null)), input))
										{
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												A.A(current);
												break;
											}
											continue;
										}
										int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
										for (int i = 1; i <= num; i = checked(i + 1))
										{
											object[] array;
											bool[] array2;
											object obj = NewLateBinding.LateGet(instance, null, VH.A(102647), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
											if (array2[0])
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
												i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
											}
											object instance2 = obj;
											if (Regex.IsMatch(Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null)), input))
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
												A.A(current);
												break;
											}
											if (Regex.IsMatch(Conversions.ToString(NewLateBinding.LateGet(instance2, null, VH.A(96399), new object[0], null, null, null)), input))
											{
												A.A(current);
												break;
											}
											instance2 = null;
										}
										instance = null;
									}
									catch (Exception ex7)
									{
										ProjectData.SetProjectError(ex7);
										Exception ex8 = ex7;
										ProjectData.ClearProjectError();
									}
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_054d;
									}
									continue;
									end_IL_054d:
									break;
								}
							}
							list = null;
							break;
						}
					}
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
			}
			if (Props.SearchForm.LookInCharts)
			{
				throw new NotImplementedException();
			}
			if (Props.SearchForm.LookInHyperlinks && range != null)
			{
				try
				{
					enumerator5 = range.GetEnumerator();
					while (enumerator5.MoveNext())
					{
						Range range5 = (Range)enumerator5.Current;
						foreach (Hyperlink hyperlink in range5.Hyperlinks)
						{
							if (Regex.IsMatch(hyperlink.Address, input))
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									A.A(range5);
									break;
								}
								break;
							}
							if (!Regex.IsMatch(hyperlink.TextToDisplay, input))
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
								A.A(range5);
								break;
							}
							break;
						}
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0666;
						}
						continue;
						end_IL_0666:
						break;
					}
				}
				finally
				{
					if (enumerator5 is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator5 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			range = null;
			return;
		}
	}

	private static object A(Range A)
	{
		return A.Text;
	}

	internal static void I(WorksheetItem A, Range B)
	{
		Text.A(A, B, H);
	}

	private static bool H(int A, int B)
	{
		return A == B;
	}

	internal static void J(WorksheetItem A, Range B)
	{
		Text.A(A, B, I);
	}

	private static bool I(int A, int B)
	{
		return A != B;
	}

	internal static void K(WorksheetItem A, Range B)
	{
		Text.A(A, B, J);
	}

	private static bool J(int A, int B)
	{
		return A > B;
	}

	internal static void L(WorksheetItem A, Range B)
	{
		Text.A(A, B, K);
	}

	private static bool K(int A, int B)
	{
		return A >= B;
	}

	internal static void M(WorksheetItem A, Range B)
	{
		Text.A(A, B, L);
	}

	private static bool L(int A, int B)
	{
		return A < B;
	}

	internal static void N(WorksheetItem A, Range B)
	{
		Text.A(A, B, M);
	}

	private static bool M(int A, int B)
	{
		return A <= B;
	}

	private static void A(WorksheetItem A, Range B, Func<int, int, bool> C)
	{
		Range range = null;
		int arg;
		try
		{
			arg = Conversions.ToInteger(Props.SearchForm.Input1);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(103879));
			ProjectData.ClearProjectError();
			return;
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
		IEnumerator enumerator6 = default(IEnumerator);
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
			if (!Props.SearchForm.LookInValues)
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
				if (!Props.SearchForm.LookInFormulas)
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
					if (!Props.SearchForm.LookInHyperlinks)
					{
						goto IL_009c;
					}
				}
			}
			range = RangeHelpers.H(B);
			goto IL_009c;
			IL_009c:
			if (Props.SearchForm.LookInValues)
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
				if (range != null)
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
					object objectValue;
					try
					{
						enumerator = range.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Range a = (Range)enumerator.Current;
							objectValue = RuntimeHelpers.GetObjectValue(Text.A(a));
							try
							{
								if (C(objectValue.ToString().Length, arg))
								{
									A.A(a);
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
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
					objectValue = null;
				}
			}
			if (Props.SearchForm.LookInFormulas)
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
				if (range != null)
				{
					{
						enumerator2 = range.GetEnumerator();
						try
						{
							while (enumerator2.MoveNext())
							{
								Range range2 = (Range)enumerator2.Current;
								try
								{
									if (!C(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null).ToString().Length, arg))
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
										A.A(range2);
										break;
									}
									continue;
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
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
									goto end_IL_01ef;
								}
								continue;
								end_IL_01ef:
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
			}
			if (Props.SearchForm.LookInComments)
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
				Range range3 = RangeHelpers.F(B);
				if (range3 != null)
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
					try
					{
						enumerator3 = range3.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							Range range4 = (Range)enumerator3.Current;
							try
							{
								if (!C(range4.Comment.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Length, arg))
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
									A.A(range4);
									break;
								}
								continue;
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_02cc;
							}
							continue;
							end_IL_02cc:
							break;
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator3 as IDisposable).Dispose();
								break;
							}
						}
					}
					range3 = null;
				}
				try
				{
					List<Range> list = RangeHelpers.A(B);
					if (list != null)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							using (List<Range>.Enumerator enumerator4 = list.GetEnumerator())
							{
								while (enumerator4.MoveNext())
								{
									Range current = enumerator4.Current;
									try
									{
										object instance = NewLateBinding.LateGet(current, null, VH.A(103833), new object[0], null, null, null);
										if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg))
										{
											while (true)
											{
												switch (7)
												{
												case 0:
													continue;
												}
												A.A(current);
												break;
											}
											continue;
										}
										if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(96399), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg))
										{
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												A.A(current);
												break;
											}
											continue;
										}
										int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
										int num2 = 1;
										while (true)
										{
											if (num2 <= num)
											{
												object[] array;
												bool[] array2;
												object obj = NewLateBinding.LateGet(instance, null, VH.A(102647), array = new object[1] { num2 }, null, null, array2 = new bool[1] { true });
												if (array2[0])
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
													num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
												}
												object instance2 = obj;
												if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg))
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
													A.A(current);
													break;
												}
												if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(96399), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg))
												{
													A.A(current);
													break;
												}
												instance2 = null;
												num2 = checked(num2 + 1);
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
											break;
										}
										instance = null;
									}
									catch (Exception ex9)
									{
										ProjectData.SetProjectError(ex9);
										Exception ex10 = ex9;
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
										goto end_IL_05f5;
									}
									continue;
									end_IL_05f5:
									break;
								}
							}
							list = null;
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
			}
			if (Props.SearchForm.LookInCharts)
			{
				throw new NotImplementedException();
			}
			if (Props.SearchForm.LookInHyperlinks)
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
				if (range != null)
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
						enumerator5 = range.GetEnumerator();
						while (enumerator5.MoveNext())
						{
							Range range5 = (Range)enumerator5.Current;
							try
							{
								enumerator6 = range5.Hyperlinks.GetEnumerator();
								while (true)
								{
									if (enumerator6.MoveNext())
									{
										Hyperlink hyperlink = (Hyperlink)enumerator6.Current;
										if (C(Conversions.ToInteger(hyperlink.Address), arg))
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												A.A(range5);
												break;
											}
											break;
										}
										if (!C(Conversions.ToInteger(hyperlink.TextToDisplay), arg))
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
											A.A(range5);
											break;
										}
										break;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_0708;
										}
										continue;
										end_IL_0708:
										break;
									}
									break;
								}
							}
							finally
							{
								if (enumerator6 is IDisposable)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										(enumerator6 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0742;
							}
							continue;
							end_IL_0742:
							break;
						}
					}
					finally
					{
						if (enumerator5 is IDisposable)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								(enumerator5 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			range = null;
			return;
		}
	}

	internal static void O(WorksheetItem A, Range B)
	{
		Text.A(A, B, (Func<int, int, int, bool>)Text.A);
	}

	private static bool A(int A, int B, int C)
	{
		if (A > B)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A < C;
				}
			}
		}
		return false;
	}

	internal static void P(WorksheetItem A, Range B)
	{
		Text.A(A, B, (Func<int, int, int, bool>)Text.B);
	}

	private static bool B(int A, int B, int C)
	{
		if (A > B)
		{
			return A >= C;
		}
		return true;
	}

	private static void A(WorksheetItem A, Range B, Func<int, int, int, bool> C)
	{
		Range range = null;
		int arg;
		int arg2;
		try
		{
			arg = Conversions.ToInteger(Props.SearchForm.Input1);
			arg2 = Conversions.ToInteger(Props.SearchForm.Input2);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(VH.A(103879));
			ProjectData.ClearProjectError();
			return;
		}
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
		IEnumerator enumerator6 = default(IEnumerator);
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
			if (!Props.SearchForm.LookInValues)
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
				if (!Props.SearchForm.LookInFormulas)
				{
					if (!Props.SearchForm.LookInHyperlinks)
					{
						goto IL_00ae;
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
				}
			}
			range = RangeHelpers.H(B);
			goto IL_00ae;
			IL_00ae:
			if (Props.SearchForm.LookInValues && range != null)
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
				object objectValue;
				{
					enumerator = range.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Range a = (Range)enumerator.Current;
							objectValue = RuntimeHelpers.GetObjectValue(Text.A(a));
							try
							{
								if (!C(objectValue.ToString().Length, arg, arg2))
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
									A.A(a);
									break;
								}
								continue;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
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
								goto end_IL_0143;
							}
							continue;
							end_IL_0143:
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
				objectValue = null;
			}
			if (Props.SearchForm.LookInFormulas)
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
				if (range != null)
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
					try
					{
						enumerator2 = range.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Range range2 = (Range)enumerator2.Current;
							try
							{
								if (!C(NewLateBinding.LateGet(range2, null, VH.A(1998), new object[0], null, null, null).ToString().Length, arg, arg2))
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
									A.A(range2);
									break;
								}
								continue;
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (7)
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
			}
			if (Props.SearchForm.LookInComments)
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
				Range range3 = RangeHelpers.F(B);
				if (range3 != null)
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
						enumerator3 = range3.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							Range range4 = (Range)enumerator3.Current;
							try
							{
								if (!C(range4.Comment.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Length, arg, arg2))
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
									A.A(range4);
									break;
								}
								continue;
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_030c;
							}
							continue;
							end_IL_030c:
							break;
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator3 as IDisposable).Dispose();
								break;
							}
						}
					}
					range3 = null;
				}
				try
				{
					List<Range> list = RangeHelpers.A(B);
					if (list != null)
					{
						using (List<Range>.Enumerator enumerator4 = list.GetEnumerator())
						{
							while (enumerator4.MoveNext())
							{
								Range current = enumerator4.Current;
								try
								{
									object instance = NewLateBinding.LateGet(current, null, VH.A(103833), new object[0], null, null, null);
									if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg, arg2))
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											A.A(current);
											break;
										}
										continue;
									}
									if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(96399), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg, arg2))
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											A.A(current);
											break;
										}
										continue;
									}
									int num = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
									int num2 = 1;
									while (true)
									{
										if (num2 <= num)
										{
											object[] array;
											bool[] array2;
											object obj = NewLateBinding.LateGet(instance, null, VH.A(102647), array = new object[1] { num2 }, null, null, array2 = new bool[1] { true });
											if (array2[0])
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
												num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
											}
											object instance2 = obj;
											if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg, arg2))
											{
												A.A(current);
												break;
											}
											if (C(Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(instance2, null, VH.A(96399), new object[0], null, null, null), null, VH.A(53246), new object[0], null, null, null)), arg, arg2))
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
												A.A(current);
												break;
											}
											instance2 = null;
											num2 = checked(num2 + 1);
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
									instance = null;
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									ProjectData.ClearProjectError();
								}
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_062a;
								}
								continue;
								end_IL_062a:
								break;
							}
						}
						list = null;
					}
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
			}
			if (Props.SearchForm.LookInCharts)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						throw new NotImplementedException();
					}
				}
			}
			if (Props.SearchForm.LookInHyperlinks)
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
				if (range != null)
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
					try
					{
						enumerator5 = range.GetEnumerator();
						while (enumerator5.MoveNext())
						{
							Range range5 = (Range)enumerator5.Current;
							{
								enumerator6 = range5.Hyperlinks.GetEnumerator();
								try
								{
									while (true)
									{
										if (enumerator6.MoveNext())
										{
											Hyperlink hyperlink = (Hyperlink)enumerator6.Current;
											if (C(Conversions.ToInteger(hyperlink.Address), arg, arg2))
											{
												while (true)
												{
													switch (5)
													{
													case 0:
														continue;
													}
													A.A(range5);
													break;
												}
												break;
											}
											if (!C(Conversions.ToInteger(hyperlink.TextToDisplay), arg, arg2))
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
												A.A(range5);
												break;
											}
											break;
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_074d;
											}
											continue;
											end_IL_074d:
											break;
										}
										break;
									}
								}
								finally
								{
									IDisposable disposable2 = enumerator6 as IDisposable;
									if (disposable2 != null)
									{
										disposable2.Dispose();
									}
								}
							}
						}
					}
					finally
					{
						if (enumerator5 is IDisposable)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								(enumerator5 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			range = null;
			return;
		}
	}
}
