using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class FastFill
{
	[CompilerGenerated]
	private static int m_A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return FastFill.m_A;
		}
		[CompilerGenerated]
		set
		{
			FastFill.m_A = value;
		}
	}

	public static void Right()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
			bool flag2 = default(bool);
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
				Application application = MH.A.Application;
				XlCalculation calculation = application.Calculation;
				if (EditMode.IsEditMode(application))
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							application = null;
							return;
						}
					}
				}
				Range range3;
				Range range;
				Range range2;
				try
				{
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					application.Calculation = XlCalculation.xlCalculationManual;
					if (Operators.ConditionalCompareObjectEqual(((Range)application.Selection).Columns.CountLarge, 1, TextCompare: false))
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
						bool flag = false;
						int num = 0;
						range = (Range)NewLateBinding.LateGet(application.Selection, null, VH.A(62391), new object[2] { 1, 1 }, null, null, null);
						int num2 = 0;
						int num3 = 10;
						while (!flag)
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
							if (num2 == num3 || range.get_Offset((object)(-num2), (object)0).Row == 1)
							{
								break;
							}
							num2++;
							range2 = range.get_Offset((object)(-num2), (object)0);
							if (Conversions.ToBoolean(Operators.NotObject(range2.EntireRow.Hidden)))
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
								if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
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
									if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)0, (object)1).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
									{
										num = range2.get_End(XlDirection.xlToRight).Column;
										flag = true;
										continue;
									}
								}
								if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0 || Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)0, (object)2).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0)
								{
									continue;
								}
								int num4 = 2;
								while (true)
								{
									range3 = range2.get_Offset((object)0, (object)num4);
									if (!Operators.ConditionalCompareObjectNotEqual(range2.NumberFormat, range3.NumberFormat, TextCompare: false))
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
										if (!string.IsNullOrEmpty(Conversions.ToString(range3.Text)))
										{
											num4 += 2;
											if (num4 <= 100)
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
											break;
										}
									}
									num = range3.get_Offset((object)0, (object)(-2)).Column;
									flag = true;
									flag2 = true;
									break;
								}
							}
							else
							{
								num3++;
							}
						}
						if (!flag)
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
								object selection = application.Selection;
								object instance = selection;
								string memberName = VH.A(62391);
								object instance2;
								object[] obj = new object[2]
								{
									NewLateBinding.LateGet(instance2 = NewLateBinding.LateGet(selection, null, VH.A(152043), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null),
									1
								};
								object[] array = obj;
								bool[] obj2 = new bool[2] { true, false };
								bool[] array2 = obj2;
								object obj3 = NewLateBinding.LateGet(instance, null, memberName, obj, null, null, obj2);
								if (array2[0])
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
									NewLateBinding.LateSetComplex(instance2, null, VH.A(152052), new object[1] { array[0] }, null, null, OptimisticSet: true, RValueBase: true);
								}
								range = (Range)obj3;
								selection = null;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							num2 = 0;
							num3 = 10;
							int num5 = Conversions.ToInteger(((Worksheet)application.ActiveSheet).Rows.CountLarge);
							while (!flag)
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
								if (num2 == num3)
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
									break;
								}
								if (range.get_Offset((object)num2, (object)0).Row != num5)
								{
									num2++;
									range2 = range.get_Offset((object)num2, (object)0);
									if (Conversions.ToBoolean(Operators.NotObject(range2.EntireRow.Hidden)))
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
										if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
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
											if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)0, (object)1).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
											{
												num = range2.get_End(XlDirection.xlToRight).Column;
												flag = true;
												continue;
											}
										}
										if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0 || Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)0, (object)2).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0)
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
										int num4 = 2;
										while (true)
										{
											range3 = range2.get_Offset((object)0, (object)num4);
											if (Operators.ConditionalCompareObjectNotEqual(range2.NumberFormat, range3.NumberFormat, TextCompare: false) || string.IsNullOrEmpty(Conversions.ToString(range3.Text)))
											{
												num = range3.get_Offset((object)0, (object)(-2)).Column;
												flag = true;
												flag2 = true;
												break;
											}
											num4 += 2;
											if (num4 <= 100)
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
											break;
										}
									}
									else
									{
										num3++;
									}
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
						}
						object selection2 = application.Selection;
						if (num > 0)
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
							((_Application)application).get_Range(RuntimeHelpers.GetObjectValue(application.Selection), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(selection2, null, VH.A(60565), new object[2]
							{
								0,
								Operators.SubtractObject(num, NewLateBinding.LateGet(selection2, null, VH.A(41354), new object[0], null, null, null))
							}, null, null, null))).Select();
						}
						selection2 = null;
					}
					JH.A((Range)application.Selection);
					Range range4 = (Range)application.Selection;
					if (Operators.ConditionalCompareObjectGreater(range4.Columns.CountLarge, 1, TextCompare: false))
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
						if (Conversions.ToBoolean(Conversions.ToBoolean(NewLateBinding.LateGet(range4.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(160260), new object[0], null, null, null)) || !A(A: false)))
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
							range4.FillRight();
							if (flag2)
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
								int num6 = Conversions.ToInteger(Operators.SubtractObject(range4.Rows.CountLarge, 1));
								for (int i = 0; i <= num6; i++)
								{
									int num7 = Conversions.ToInteger(range4.Columns.CountLarge);
									for (int num4 = 1; num4 <= num7; num4 += 2)
									{
										object[] array;
										bool[] array2;
										object instance3 = NewLateBinding.LateGet(range4.Cells[1, 1], null, VH.A(60565), array = new object[2] { i, num4 }, null, null, array2 = new bool[2] { true, true });
										if (array2[0])
										{
											i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
										}
										if (array2[1])
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
											num4 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(int));
										}
										NewLateBinding.LateCall(instance3, null, VH.A(160281), new object[0], null, null, null, IgnoreReturn: true);
										NewLateBinding.LateCall(instance3, null, VH.A(160308), new object[0], null, null, null, IgnoreReturn: true);
										NewLateBinding.LateCall(instance3, null, VH.A(160333), new object[0], null, null, null, IgnoreReturn: true);
										_ = null;
									}
								}
							}
						}
					}
					range4 = null;
					JH.A((Range)application.Selection, VH.A(160360));
					Common.RefreshLiveVisualizations((Range)application.Selection);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				application.Calculation = calculation;
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				application = null;
				range3 = null;
				range = null;
				range2 = null;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(160360));
				return;
			}
		}
	}

	public static void Down()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
			bool flag2 = default(bool);
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
				Application application = MH.A.Application;
				XlCalculation calculation = application.Calculation;
				if (EditMode.IsEditMode(application))
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							application = null;
							return;
						}
					}
				}
				Range range3;
				Range range;
				Range range2;
				try
				{
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					application.Calculation = XlCalculation.xlCalculationManual;
					if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(application.Selection, null, VH.A(152043), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null), 1, TextCompare: false))
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
						bool flag = false;
						int num = 0;
						range = (Range)NewLateBinding.LateGet(application.Selection, null, VH.A(62391), new object[2] { 1, 1 }, null, null, null);
						int num2 = 0;
						int num3 = 10;
						while (!flag)
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
							if (num2 == num3)
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
								break;
							}
							if (range.get_Offset((object)0, (object)(-num2)).Column != 1)
							{
								num2++;
								range2 = range.get_Offset((object)0, (object)(-num2));
								if (Conversions.ToBoolean(Operators.NotObject(range2.EntireColumn.Hidden)))
								{
									if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
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
										if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)1, (object)0).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
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
											num = range2.get_End(XlDirection.xlDown).Row;
											flag = true;
											continue;
										}
									}
									if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0)
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
									if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)2, (object)0).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0)
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
									int num4 = 2;
									while (true)
									{
										range3 = range2.get_Offset((object)num4, (object)0);
										if (!Operators.ConditionalCompareObjectNotEqual(range2.NumberFormat, range3.NumberFormat, TextCompare: false))
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
											if (!string.IsNullOrEmpty(Conversions.ToString(range3.Text)))
											{
												num4 += 2;
												if (num4 <= 100)
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
										}
										num = range3.get_Offset((object)(-2), (object)0).Row;
										flag = true;
										flag2 = true;
										break;
									}
								}
								else
								{
									num3++;
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
						if (!flag)
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
								Range range4 = (Range)application.Selection;
								range = (Range)range4.Cells[1, RuntimeHelpers.GetObjectValue(range4.Columns.CountLarge)];
								range4 = null;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							num2 = 0;
							num3 = 10;
							int num5 = Conversions.ToInteger(((Worksheet)application.ActiveSheet).Columns.CountLarge);
							while (!flag)
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
								if (num2 == num3)
								{
									break;
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
								if (range.get_Offset((object)0, (object)num2).Column != num5)
								{
									num2++;
									range2 = range.get_Offset((object)0, (object)num2);
									if (Conversions.ToBoolean(Operators.NotObject(range2.EntireColumn.Hidden)))
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
										if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
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
											if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)1, (object)0).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) != 0)
											{
												num = range2.get_End(XlDirection.xlDown).Row;
												flag = true;
												continue;
											}
										}
										if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0)
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
										if (Operators.CompareString(Strings.Trim(Conversions.ToString(range2.get_Offset((object)2, (object)0).get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), "", TextCompare: false) == 0)
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
										int num4 = 2;
										while (true)
										{
											range3 = range2.get_Offset((object)num4, (object)0);
											if (!Operators.ConditionalCompareObjectNotEqual(range2.NumberFormat, range3.NumberFormat, TextCompare: false))
											{
												if (!string.IsNullOrEmpty(Conversions.ToString(range3.Text)))
												{
													num4 += 2;
													if (num4 <= 100)
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
													break;
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
											num = range3.get_Offset((object)(-2), (object)0).Row;
											flag = true;
											flag2 = true;
											break;
										}
									}
									else
									{
										num3++;
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
						}
						object selection = application.Selection;
						if (num > 0)
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
							((_Application)application).get_Range(RuntimeHelpers.GetObjectValue(application.Selection), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(selection, null, VH.A(60565), new object[2]
							{
								Operators.SubtractObject(num, NewLateBinding.LateGet(selection, null, VH.A(91823), new object[0], null, null, null)),
								0
							}, null, null, null))).Select();
						}
						selection = null;
					}
					JH.A((Range)application.Selection);
					Range range5 = (Range)application.Selection;
					if (Operators.ConditionalCompareObjectGreater(range5.Rows.CountLarge, 1, TextCompare: false))
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
						if (Conversions.ToBoolean(Conversions.ToBoolean(NewLateBinding.LateGet(range5.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(160260), new object[0], null, null, null)) || !A(A: true)))
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
							range5.FillDown();
							if (flag2)
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
								int num6 = Conversions.ToInteger(Operators.SubtractObject(range5.Columns.CountLarge, 1));
								for (int i = 0; i <= num6; i++)
								{
									int num7 = Conversions.ToInteger(range5.Rows.CountLarge);
									for (int num4 = 1; num4 <= num7; num4 += 2)
									{
										object[] array;
										bool[] array2;
										object instance = NewLateBinding.LateGet(range5.Cells[1, 1], null, VH.A(60565), array = new object[2] { num4, i }, null, null, array2 = new bool[2] { true, true });
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
											num4 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
										}
										if (array2[1])
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
											i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(int));
										}
										NewLateBinding.LateCall(instance, null, VH.A(160281), new object[0], null, null, null, IgnoreReturn: true);
										NewLateBinding.LateCall(instance, null, VH.A(160308), new object[0], null, null, null, IgnoreReturn: true);
										NewLateBinding.LateCall(instance, null, VH.A(160333), new object[0], null, null, null, IgnoreReturn: true);
										_ = null;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_091e;
										}
										continue;
										end_IL_091e:
										break;
									}
								}
							}
						}
					}
					range5 = null;
					JH.A((Range)application.Selection, VH.A(160391));
					Common.RefreshLiveVisualizations((Range)application.Selection);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				application.Calculation = calculation;
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				application = null;
				range3 = null;
				range = null;
				range2 = null;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(160391));
				return;
			}
		}
	}

	private static bool A(bool A)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		Application application = default(Application);
		int num = default(int);
		int num3 = default(int);
		string text = default(string);
		Range range = default(Range);
		int num5 = default(int);
		int num6 = default(int);
		int num7 = default(int);
		int num8 = default(int);
		string text2 = default(string);
		string text3 = default(string);
		string text4 = default(string);
		string[] array = default(string[]);
		int num9 = default(int);
		bool result = default(bool);
		XlAutoFillType xlAutoFillType = default(XlAutoFillType);
		bool flag = default(bool);
		MatchCollection matchCollection = default(MatchCollection);
		Match match = default(Match);
		MatchCollection matchCollection2 = default(MatchCollection);
		Range range2 = default(Range);
		Match match2 = default(Match);
		MatchCollection matchCollection3 = default(MatchCollection);
		Match match3 = default(Match);
		string text5 = default(string);
		int num10 = default(int);
		int num11 = default(int);
		string text6 = default(string);
		int num12 = default(int);
		int num13 = default(int);
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					object instance;
					string memberName;
					object[] obj;
					object[] array2;
					bool[] obj2;
					bool[] array3;
					int result2;
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						application = MH.A.Application;
						goto IL_0012;
					case 3618:
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
								goto IL_0012;
							case 3:
								goto IL_0017;
							case 4:
								goto IL_001c;
							case 5:
								goto IL_0021;
							case 6:
								goto IL_003f;
							case 7:
								goto IL_0048;
							case 8:
								goto IL_004a;
							case 9:
								goto IL_0053;
							case 10:
								goto IL_005d;
							case 11:
								goto IL_0067;
							case 12:
								goto IL_006e;
							case 13:
								goto IL_0081;
							case 14:
								goto IL_00dd;
							case 15:
								goto IL_00ed;
							case 17:
								goto IL_00fc;
							case 16:
							case 18:
								goto IL_0109;
							case 19:
								goto IL_0113;
							case 20:
								goto IL_0170;
							case 21:
								goto IL_017f;
							case 23:
								goto IL_01a6;
							case 25:
								goto IL_01ae;
							case 27:
								goto IL_01b6;
							case 29:
								goto IL_01be;
							case 30:
								goto IL_01c4;
							case 22:
							case 24:
							case 26:
							case 28:
							case 31:
							case 32:
								goto IL_01cd;
							case 33:
								goto IL_0289;
							case 34:
								goto IL_0293;
							case 35:
								goto IL_02d8;
							case 36:
								goto IL_02f3;
							case 37:
								goto IL_030e;
							case 38:
								goto IL_032a;
							case 39:
								goto IL_0339;
							case 40:
								goto IL_0350;
							case 41:
								goto IL_036d;
							case 42:
								goto IL_038f;
							case 44:
								goto IL_03af;
							case 45:
								goto IL_03b2;
							case 46:
								goto IL_03b5;
							case 47:
								goto IL_03d0;
							case 48:
								goto IL_03ec;
							case 49:
								goto IL_0400;
							case 50:
								goto IL_040f;
							case 51:
								goto IL_0431;
							case 52:
								goto IL_044c;
							case 53:
								goto IL_0469;
							case 55:
								goto IL_0471;
							case 56:
								goto IL_0474;
							case 57:
								goto IL_048f;
							case 58:
								goto IL_04ab;
							case 59:
								goto IL_04ba;
							case 60:
								goto IL_04d1;
							case 61:
								goto IL_04e0;
							case 62:
								goto IL_0500;
							case 63:
								goto IL_051d;
							case 64:
								goto IL_0538;
							case 66:
								goto IL_0540;
							case 67:
								goto IL_0543;
							case 68:
								goto IL_0552;
							case 69:
								goto IL_058d;
							case 70:
								goto IL_05a9;
							case 71:
								goto IL_05bd;
							case 73:
								goto IL_05ec;
							case 75:
								goto IL_0600;
							case 76:
								goto IL_0611;
							case 77:
								goto IL_064a;
							case 78:
								goto IL_0664;
							case 79:
								goto IL_0678;
							case 81:
								goto IL_06a7;
							case 83:
								goto IL_06c5;
							case 84:
								goto IL_06d4;
							case 85:
								goto IL_070c;
							case 86:
								goto IL_0726;
							case 87:
								goto IL_073a;
							case 89:
								goto IL_0767;
							case 91:
								goto IL_0785;
							case 92:
								goto IL_0796;
							case 93:
								goto IL_07ce;
							case 94:
								goto IL_07ea;
							case 95:
								goto IL_07fe;
							case 97:
								goto IL_082a;
							case 99:
								goto IL_0845;
							case 43:
							case 54:
							case 65:
							case 72:
							case 74:
							case 80:
							case 82:
							case 88:
							case 90:
							case 96:
							case 98:
							case 100:
								goto IL_084f;
							case 101:
								goto IL_0872;
							case 102:
								goto IL_0887;
							case 103:
								goto IL_08d1;
							case 104:
								goto IL_08ed;
							case 105:
								goto IL_08f9;
							case 106:
								goto IL_0923;
							case 107:
								goto IL_0929;
							case 109:
								goto IL_0950;
							case 110:
								goto IL_0959;
							case 111:
								goto IL_0988;
							case 112:
								goto IL_0997;
							case 114:
								goto IL_09b5;
							case 115:
								goto IL_09be;
							case 116:
								goto IL_09cf;
							case 118:
								goto IL_09ed;
							case 120:
								goto IL_0a26;
							case 121:
								goto IL_0a2c;
							case 122:
								goto IL_0a3d;
							case 124:
								goto IL_0a56;
							case 125:
								goto IL_0a5f;
							case 108:
							case 113:
							case 117:
							case 119:
							case 123:
							case 126:
							case 127:
								goto IL_0a96;
							case 129:
								goto IL_0aa6;
							case 130:
								goto IL_0ab2;
							case 131:
								goto IL_0ac4;
							case 132:
								goto IL_0acd;
							case 133:
								goto IL_0ada;
							case 134:
								goto IL_0aee;
							case 128:
							case 135:
								goto IL_0afd;
							case 136:
								goto IL_0b18;
							case 137:
								goto IL_0b24;
							case 138:
								goto IL_0b42;
							case 139:
								goto IL_0b5d;
							case 140:
								goto IL_0b80;
							case 141:
								goto IL_0bae;
							case 142:
								goto IL_0bbc;
							case 143:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 144:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_0887:
						num2 = 102;
						text = Conversions.ToString(NewLateBinding.LateGet(range.Cells[num5 - 1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(96399), new object[0], null, null, null));
						goto IL_08d1;
						IL_0012:
						num2 = 2;
						num6 = 0;
						goto IL_0017;
						IL_0017:
						num2 = 3;
						num7 = 0;
						goto IL_001c;
						IL_001c:
						num2 = 4;
						num5 = 1;
						goto IL_0021;
						IL_0021:
						num2 = 5;
						num8 = Conversions.ToInteger(application.Evaluate(VH.A(160420)));
						goto IL_003f;
						IL_003f:
						num2 = 6;
						text = "";
						goto IL_0048;
						IL_0048:
						num2 = 7;
						goto IL_004a;
						IL_004a:
						num2 = 8;
						text2 = "";
						goto IL_0053;
						IL_0053:
						num2 = 9;
						text3 = "";
						goto IL_005d;
						IL_005d:
						num2 = 10;
						text4 = "";
						goto IL_0067;
						IL_0067:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_006e;
						IL_006e:
						num2 = 12;
						range = (Range)application.Selection;
						goto IL_0081;
						IL_0081:
						num2 = 13;
						if (Operators.ConditionalCompareObjectGreater(range.Rows.CountLarge, 1, TextCompare: false))
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
							if (Operators.ConditionalCompareObjectGreater(range.Columns.CountLarge, 1, TextCompare: false))
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
								goto IL_00dd;
							}
						}
						goto IL_0113;
						IL_08d1:
						num2 = 103;
						if (Information.UBound(array) == 0)
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
							goto IL_08ed;
						}
						goto IL_0aa6;
						IL_0ac4:
						num2 = 131;
						num9 = 0;
						goto IL_0acd;
						IL_08f9:
						num2 = 105;
						if (unchecked(Strings.InStr(text, VH.A(150436)) > 0 && num6 > 1))
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
							goto IL_0923;
						}
						goto IL_0929;
						IL_08ed:
						num2 = 104;
						text = Strings.UCase(text);
						goto IL_08f9;
						IL_0acd:
						num2 = 132;
						text = array[num9];
						goto IL_0ada;
						IL_00dd:
						num2 = 14;
						if (A)
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
							goto IL_00ed;
						}
						goto IL_00fc;
						IL_0923:
						num2 = 106;
						num6 = 4;
						goto IL_0929;
						IL_00ed:
						num2 = 15;
						range.FillDown();
						goto IL_0109;
						IL_00fc:
						num2 = 17;
						range.FillRight();
						goto IL_0109;
						IL_0109:
						num2 = 18;
						result = true;
						goto end_IL_0000_3;
						IL_0113:
						num2 = 19;
						if (Information.IsDate(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41636), new object[0], null, null, null))))
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
							goto IL_0170;
						}
						goto IL_0293;
						IL_0929:
						num2 = 107;
						if (num6 == 0)
						{
							goto IL_0950;
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
						if (unchecked((uint)(num6 - 1)) <= 2u)
						{
							goto IL_09b5;
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
						goto IL_0a26;
						IL_0170:
						num2 = 20;
						CycleIndex++;
						goto IL_017f;
						IL_017f:
						num2 = 21;
						switch (CycleIndex)
						{
						case 1:
							break;
						case 2:
							goto IL_01ae;
						case 3:
							goto IL_01b6;
						case 4:
							goto IL_01be;
						default:
							goto IL_01cd;
						}
						goto IL_01a6;
						IL_01be:
						num2 = 29;
						xlAutoFillType = XlAutoFillType.xlFillYears;
						goto IL_01c4;
						IL_01c4:
						num2 = 30;
						CycleIndex = 0;
						goto IL_01cd;
						IL_01b6:
						num2 = 27;
						xlAutoFillType = XlAutoFillType.xlFillMonths;
						goto IL_01cd;
						IL_01ae:
						num2 = 25;
						xlAutoFillType = XlAutoFillType.xlFillWeekdays;
						goto IL_01cd;
						IL_01a6:
						num2 = 23;
						xlAutoFillType = XlAutoFillType.xlFillDays;
						goto IL_01cd;
						IL_01cd:
						num2 = 32;
						instance = range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
						memberName = VH.A(160445);
						obj = new object[2] { range, xlAutoFillType };
						array2 = obj;
						obj2 = new bool[2] { true, true };
						array3 = obj2;
						NewLateBinding.LateCall(instance, null, memberName, obj, null, null, obj2, IgnoreReturn: true);
						if (array3[0])
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
							range = (Range)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(Range));
						}
						if (array3[1])
						{
							xlAutoFillType = (XlAutoFillType)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(XlAutoFillType));
						}
						goto IL_0289;
						IL_0845:
						num2 = 99;
						result = false;
						goto end_IL_0000_3;
						IL_0aee:
						num2 = 134;
						text = Strings.UCase(text);
						goto IL_0afd;
						IL_0ada:
						num2 = 133;
						if (flag)
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
							goto IL_0aee;
						}
						goto IL_0afd;
						IL_0a26:
						num2 = 120;
						num6 = 1;
						goto IL_0a2c;
						IL_0289:
						num2 = 33;
						result = true;
						goto end_IL_0000_3;
						IL_0293:
						num2 = 34;
						text = Conversions.ToString(NewLateBinding.LateGet(range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(96399), new object[0], null, null, null));
						goto IL_02d8;
						IL_02d8:
						num2 = 35;
						flag = Regex.IsMatch(text, VH.A(160462), RegexOptions.None);
						goto IL_02f3;
						IL_02f3:
						num2 = 36;
						matchCollection = Regex.Matches(text, VH.A(160479), RegexOptions.None);
						goto IL_030e;
						IL_030e:
						num2 = 37;
						if (matchCollection.Count > 0)
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
							goto IL_032a;
						}
						goto IL_03b2;
						IL_0a2c:
						num2 = 121;
						if (num7 == 0)
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
							goto IL_0a3d;
						}
						goto IL_0a56;
						IL_032a:
						num2 = 38;
						match = matchCollection[0];
						goto IL_0339;
						IL_0339:
						num2 = 39;
						if (!int.TryParse(match.Value, out result2))
						{
							goto IL_0350;
						}
						goto IL_03af;
						IL_0350:
						num2 = 40;
						text2 = match.Groups[1].ToString();
						goto IL_036d;
						IL_036d:
						num2 = 41;
						num7 = Conversions.ToInteger(match.Groups[3].Value);
						goto IL_038f;
						IL_038f:
						num2 = 42;
						text3 = match.Groups[4].ToString();
						goto IL_084f;
						IL_03af:
						match = null;
						goto IL_03b2;
						IL_03b2:
						matchCollection = null;
						goto IL_03b5;
						IL_03b5:
						num2 = 46;
						matchCollection2 = Regex.Matches(text, VH.A(160556), RegexOptions.None);
						goto IL_03d0;
						IL_03d0:
						num2 = 47;
						if (matchCollection2.Count > 0)
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
							goto IL_03ec;
						}
						goto IL_0471;
						IL_0afd:
						num2 = 135;
						range2.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)text);
						goto IL_0b18;
						IL_03ec:
						num2 = 48;
						num6 = Conversions.ToInteger(Strings.Left(text, 1));
						goto IL_0400;
						IL_0400:
						num2 = 49;
						match2 = matchCollection2[0];
						goto IL_040f;
						IL_040f:
						num2 = 50;
						num7 = Conversions.ToInteger(match2.Groups[2].ToString());
						goto IL_0431;
						IL_0431:
						num2 = 51;
						text4 = match2.Groups[1].ToString();
						goto IL_044c;
						IL_044c:
						num2 = 52;
						text3 = match2.Groups[3].ToString();
						goto IL_0469;
						IL_0469:
						match2 = null;
						goto IL_084f;
						IL_0471:
						matchCollection2 = null;
						goto IL_0474;
						IL_0474:
						num2 = 56;
						matchCollection3 = Regex.Matches(text, VH.A(160629), RegexOptions.None);
						goto IL_048f;
						IL_048f:
						num2 = 57;
						if (matchCollection3.Count > 0)
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
							goto IL_04ab;
						}
						goto IL_0540;
						IL_0a3d:
						num2 = 122;
						text = text2 + Conversions.ToString(num6) + text4;
						goto IL_0a96;
						IL_04ab:
						num2 = 58;
						text2 = Strings.Left(text, 1);
						goto IL_04ba;
						IL_04ba:
						num2 = 59;
						num6 = Conversions.ToInteger(Strings.Mid(text, 2, 1));
						goto IL_04d1;
						IL_04d1:
						num2 = 60;
						match3 = matchCollection3[0];
						goto IL_04e0;
						IL_04e0:
						num2 = 61;
						num7 = Conversions.ToInteger(match3.Groups[2].Value);
						goto IL_0500;
						IL_0500:
						num2 = 62;
						text4 = match3.Groups[1].ToString();
						goto IL_051d;
						IL_051d:
						num2 = 63;
						text3 = match3.Groups[3].ToString();
						goto IL_0538;
						IL_0538:
						match3 = null;
						goto IL_084f;
						IL_0540:
						matchCollection3 = null;
						goto IL_0543;
						IL_0543:
						num2 = 67;
						text5 = VH.A(160702);
						goto IL_0552;
						IL_0552:
						num2 = 68;
						if (Regex.IsMatch(text, VH.A(160797) + text5 + VH.A(160802), RegexOptions.None))
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
							goto IL_058d;
						}
						goto IL_0600;
						IL_0a56:
						num2 = 124;
						num7++;
						goto IL_0a5f;
						IL_058d:
						num2 = 69;
						array = Strings.Split(text5, VH.A(43203));
						goto IL_05a9;
						IL_05a9:
						num2 = 70;
						num10 = Information.UBound(array);
						num9 = 0;
						goto IL_05f5;
						IL_05f5:
						if (num9 <= num10)
						{
							goto IL_05bd;
						}
						goto IL_084f;
						IL_05bd:
						num2 = 71;
						if (Operators.CompareString(Strings.UCase(text), Strings.UCase(array[num9]), TextCompare: false) != 0)
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
							goto IL_05ec;
						}
						goto IL_084f;
						IL_0a5f:
						num2 = 125;
						text = text2 + Conversions.ToString(num6) + text4 + FastFill.A(num7) + text3;
						goto IL_0a96;
						IL_05ec:
						num2 = 73;
						num9++;
						goto IL_05f5;
						IL_0600:
						num2 = 75;
						text5 = VH.A(160807);
						goto IL_0611;
						IL_0611:
						num2 = 76;
						if (Regex.IsMatch(text, VH.A(160797) + text5 + VH.A(160802), RegexOptions.None))
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
							goto IL_064a;
						}
						goto IL_06c5;
						IL_09b5:
						num2 = 114;
						num6++;
						goto IL_09be;
						IL_064a:
						num2 = 77;
						array = Strings.Split(text5, VH.A(43203));
						goto IL_0664;
						IL_0664:
						num2 = 78;
						num11 = Information.UBound(array);
						num9 = 0;
						goto IL_06b0;
						IL_06b0:
						if (num9 <= num11)
						{
							goto IL_0678;
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
						goto IL_084f;
						IL_09be:
						num2 = 115;
						if (num7 == 0)
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
							goto IL_09cf;
						}
						goto IL_09ed;
						IL_0678:
						num2 = 79;
						if (Operators.CompareString(Strings.UCase(text), Strings.UCase(array[num9]), TextCompare: false) != 0)
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
							goto IL_06a7;
						}
						goto IL_084f;
						IL_0b18:
						num2 = 136;
						num5++;
						goto IL_0b24;
						IL_06a7:
						num2 = 81;
						num9++;
						goto IL_06b0;
						IL_06c5:
						num2 = 83;
						text6 = VH.A(160979);
						goto IL_06d4;
						IL_06d4:
						num2 = 84;
						if (Regex.IsMatch(text, VH.A(160797) + text6 + VH.A(160802)))
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
							goto IL_070c;
						}
						goto IL_0785;
						IL_09cf:
						num2 = 116;
						text = text2 + Conversions.ToString(num6) + text4;
						goto IL_0a96;
						IL_070c:
						num2 = 85;
						array = Strings.Split(text6, VH.A(43203));
						goto IL_0726;
						IL_0726:
						num2 = 86;
						num12 = Information.UBound(array);
						num9 = 0;
						goto IL_0770;
						IL_0770:
						if (num9 <= num12)
						{
							goto IL_073a;
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
						goto IL_084f;
						IL_09ed:
						num2 = 118;
						text = text2 + Conversions.ToString(num6) + text4 + FastFill.A(num7) + text3;
						goto IL_0a96;
						IL_073a:
						num2 = 87;
						if (Operators.CompareString(Strings.UCase(text), Strings.UCase(array[num9]), TextCompare: false) != 0)
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
							goto IL_0767;
						}
						goto IL_084f;
						IL_0950:
						num2 = 109;
						num7++;
						goto IL_0959;
						IL_0767:
						num2 = 89;
						num9++;
						goto IL_0770;
						IL_0785:
						num2 = 91;
						text6 = VH.A(161034);
						goto IL_0796;
						IL_0796:
						num2 = 92;
						if (Regex.IsMatch(text, VH.A(160797) + text6 + VH.A(160802)))
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
							goto IL_07ce;
						}
						goto IL_0845;
						IL_0959:
						num2 = 110;
						if ((Operators.CompareString(text3, VH.A(57237), TextCompare: false) == 0) & (num7 > num8 - 1))
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
							goto IL_0988;
						}
						goto IL_0997;
						IL_07ce:
						num2 = 93;
						array = Strings.Split(text6, VH.A(43203));
						goto IL_07ea;
						IL_07ea:
						num2 = 94;
						num13 = Information.UBound(array);
						num9 = 0;
						goto IL_0833;
						IL_0833:
						if (num9 <= num13)
						{
							goto IL_07fe;
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
						goto IL_084f;
						IL_0b24:
						num2 = 137;
						goto IL_0b2a;
						IL_07fe:
						num2 = 95;
						if (Operators.CompareString(Strings.UCase(text), Strings.UCase(array[num9]), TextCompare: false) != 0)
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
							goto IL_082a;
						}
						goto IL_084f;
						IL_0988:
						num2 = 111;
						text3 = VH.A(150457);
						goto IL_0997;
						IL_082a:
						num2 = 97;
						num9++;
						goto IL_0833;
						IL_084f:
						num2 = 100;
						enumerator = range.GetEnumerator();
						goto IL_0b2a;
						IL_0b2a:
						if (enumerator.MoveNext())
						{
							range2 = (Range)enumerator.Current;
							goto IL_0872;
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
						goto IL_0b42;
						IL_0997:
						num2 = 112;
						text = text2 + Conversions.ToString(num7) + text3;
						goto IL_0a96;
						IL_0b42:
						num2 = 138;
						if (enumerator is IDisposable)
						{
							(enumerator as IDisposable).Dispose();
						}
						goto IL_0b5d;
						IL_0aa6:
						num2 = 129;
						num9++;
						goto IL_0ab2;
						IL_0b5d:
						num2 = 139;
						application.ActiveCell.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0b80;
						IL_0b80:
						num2 = 140;
						range.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0bae;
						IL_0bae:
						num2 = 141;
						application.CutCopyMode = (XlCutCopyMode)0;
						goto IL_0bbc;
						IL_0bbc:
						num2 = 142;
						range = null;
						break;
						IL_0a96:
						num2 = 127;
						text = Strings.UCase(text);
						goto IL_0afd;
						IL_0872:
						num2 = 101;
						if (num5 > 1)
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
							goto IL_0887;
						}
						goto IL_0b18;
						IL_0ab2:
						num2 = 130;
						if (num9 > Information.UBound(array))
						{
							goto IL_0ac4;
						}
						goto IL_0acd;
						end_IL_0000_2:
						break;
					}
					num2 = 143;
					result = true;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj3) when (obj3 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj3);
				try0000_dispatch = 3618;
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
				switch (1)
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

	private static string A(int A)
	{
		if (A < 10)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return VH.A(64024) + Conversions.ToString(A);
				}
			}
		}
		return Conversions.ToString(A);
	}
}
