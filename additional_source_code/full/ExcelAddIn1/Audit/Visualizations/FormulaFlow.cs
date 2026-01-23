using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Visualizations;

public sealed class FormulaFlow
{
	[CompilerGenerated]
	internal sealed class DD
	{
		public Range A;

		[SpecialName]
		internal bool A(Range A)
		{
			return Operators.CompareString(A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), this.A.get_Offset((object)0, (object)(-1)).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0;
		}

		[SpecialName]
		internal bool B(Range A)
		{
			return Operators.CompareString(A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), this.A.get_Offset((object)(-1), (object)0).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	private static int m_A = ColorTranslator.ToOle(Color.FromArgb(255, 0, 200));

	private static int PATTERN_COLOR_VIOLET
	{
		[CompilerGenerated]
		get
		{
			return FormulaFlow.m_A;
		}
	}

	public static void Apply()
	{
		if (!Licensing.AllowVisualizationOperation())
		{
			return;
		}
		int B = default(int);
		int C = default(int);
		int D = default(int);
		int E = default(int);
		XlCalculation calculation = default(XlCalculation);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			Microsoft.Office.Interop.Excel.Window window = null;
			if (!(application.Selection is Range))
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						application = null;
						return;
					}
				}
			}
			Range range = JH.A((Range)null);
			if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
				range = range.CurrentRegion;
			}
			Range range2 = ExcelAddIn1.Formulas.Helpers.SpecialCellsFormulas(range);
			if (range2 != null)
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
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
				Worksheet worksheet;
				Worksheet worksheet2;
				if (!Workbooks.IsShared(activeWorkbook, true, (System.Windows.Window)null))
				{
					List<Range> H = new List<Range>();
					List<Range> I = new List<Range>();
					Common.A(range, Common.VisualizationType.FormulaFlow);
					FormulaFlow.B(application);
					A(application);
					A(range, ref B, ref C, ref D, ref E);
					if (activeWorkbook.Windows.Count > 1)
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
						window = application.ActiveWindow;
					}
					try
					{
						Microsoft.Office.Interop.Excel.Application application2 = application;
						application2.ScreenUpdating = false;
						application2.EnableEvents = false;
						application2.DisplayAlerts = false;
						calculation = application2.Calculation;
						application2.Calculation = XlCalculation.xlCalculationManual;
						NewLateBinding.LateCall(application2.ActiveSheet, null, VH.A(224), new object[1] { application2.ActiveSheet }, new string[1] { VH.A(51175) }, null, null, IgnoreReturn: true);
						worksheet = (Worksheet)application2.ActiveSheet;
						application2 = null;
						if (window != null)
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
							window.Activate();
							window = null;
						}
						worksheet2 = range.Worksheet;
						foreach (Range item in range2)
						{
							FormulaFlow.C(item);
							FormulaFlow.B(item, worksheet, worksheet2, B, C, D, E, ref H, ref I);
						}
						worksheet.Delete();
						range.Select();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					Microsoft.Office.Interop.Excel.Application application3 = application;
					application3.ScreenUpdating = true;
					application3.EnableEvents = true;
					application3.DisplayAlerts = true;
					application3.Calculation = calculation;
					_ = null;
					H = null;
					I = null;
				}
				range2 = null;
				worksheet = null;
				worksheet2 = null;
				activeWorkbook = null;
			}
			application = null;
			Common.A(VH.A(51188));
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.Excel.Application A)
	{
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(51213)).AddEventHandler(A, new AppEvents_SheetChangeEventHandler(FormulaFlow.A));
	}

	internal static void B(Microsoft.Office.Interop.Excel.Application A)
	{
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(51213)).RemoveEventHandler(A, new AppEvents_SheetChangeEventHandler(FormulaFlow.A));
	}

	internal static void A(object A, Range B)
	{
		Microsoft.Office.Interop.Excel.Application application = B.Application;
		Worksheet worksheet = (Worksheet)A;
		Microsoft.Office.Interop.Excel.Window window = null;
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
		checked
		{
			if (Operators.CompareString(((Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent).Name, activeWorkbook.Name, TextCompare: false) == 0)
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
				if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(B.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51236), new object[0], null, null, null), null, VH.A(5814), new object[0], null, null, null), B.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) && Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(B.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51255), new object[0], null, null, null), null, VH.A(5814), new object[0], null, null, null), B.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
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
					List<Range> H = new List<Range>();
					List<Range> I = new List<Range>();
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = B.Cells.GetEnumerator();
						while (enumerator.MoveNext())
						{
							FormulaFlow.C((Range)enumerator.Current);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_01d1;
							}
							continue;
							end_IL_01d1:
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
					int B2 = default(int);
					int C = default(int);
					int D = default(int);
					int E = default(int);
					IEnumerator enumerator2 = default(IEnumerator);
					IEnumerator enumerator3 = default(IEnumerator);
					Worksheet worksheet2;
					Range key;
					Range arg;
					Range range3;
					for (int i = Common.VisualizedRanges.Count - 1; i >= 0; i += -1)
					{
						key = Common.VisualizedRanges[i].Key;
						try
						{
							key.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							Common.VisualizedRanges.RemoveAt(i);
							ProjectData.ClearProjectError();
							continue;
						}
						if (Common.VisualizedRanges[i].Value != Common.VisualizationType.FormulaFlow || application.Intersect(key, B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
						{
							continue;
						}
						FormulaFlow.A(key, ref B2, ref C, ref D, ref E);
						if (activeWorkbook.Windows.Count > 1)
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
							window = application.ActiveWindow;
						}
						Microsoft.Office.Interop.Excel.Application application2 = application;
						application2.ScreenUpdating = false;
						application2.EnableEvents = false;
						application2.DisplayAlerts = false;
						XlCalculation calculation = application2.Calculation;
						application2.Calculation = XlCalculation.xlCalculationManual;
						NewLateBinding.LateCall(application2.ActiveSheet, null, VH.A(224), new object[1] { application2.ActiveSheet }, new string[1] { VH.A(51175) }, null, null, IgnoreReturn: true);
						worksheet2 = (Worksheet)application2.ActiveSheet;
						application2 = null;
						if (window != null)
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
							window.Activate();
							window = null;
						}
						try
						{
							enumerator2 = B.Cells.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Range range = (Range)enumerator2.Current;
								if (application.Intersect(key, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
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
								try
								{
									Range range2 = range;
									if (range2.Column > 1)
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
										if (range2.Row > 1)
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
											arg = range2.get_Offset((object)(-1), (object)(-1)).get_Resize((object)3, (object)3);
											goto IL_0747;
										}
									}
									if (range2.Column > 1)
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
										arg = range2.get_Offset((object)0, (object)(-1)).get_Resize((object)2, (object)3);
									}
									else if (range2.Row > 1)
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
										arg = range2.get_Offset((object)(-1), (object)0).get_Resize((object)3, (object)2);
									}
									else
									{
										arg = range2.get_Resize((object)2, (object)2);
									}
									goto IL_0747;
									IL_0747:
									range2 = null;
									range3 = ExcelAddIn1.Formulas.Helpers.SpecialCellsFormulas(application.Intersect(key, arg, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
									try
									{
										enumerator3 = range3.GetEnumerator();
										while (enumerator3.MoveNext())
										{
											FormulaFlow.B((Range)enumerator3.Current, worksheet2, worksheet, B2, C, D, E, ref H, ref I);
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_08d5;
											}
											continue;
											end_IL_08d5:
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
								switch (5)
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
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						worksheet2.Delete();
						Microsoft.Office.Interop.Excel.Application application3 = application;
						application3.Calculation = calculation;
						application3.EnableEvents = true;
						application3.ScreenUpdating = true;
						application3.DisplayAlerts = true;
						_ = null;
						break;
					}
					H = null;
					I = null;
					worksheet2 = null;
					key = null;
					arg = null;
					range3 = null;
				}
			}
			application = null;
			worksheet = null;
			activeWorkbook = null;
		}
	}

	private static void A(Range A, ref int B, ref int C, ref int D, ref int E)
	{
		Range range = A;
		B = ((Range)range.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Column;
		C = ((Range)range.Columns[RuntimeHelpers.GetObjectValue(range.Columns.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).Column;
		D = ((Range)range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Row;
		E = ((Range)range.Rows[RuntimeHelpers.GetObjectValue(range.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).Row;
		range = null;
	}

	private static void B(Range A, Worksheet B, Worksheet C, int D, int E, int F, int G, ref List<Range> H, ref List<Range> I)
	{
		bool flag = false;
		bool flag2 = false;
		Range a = ((_Worksheet)B).get_Range((object)A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value));
		if (Conversions.ToBoolean(A.HasArray))
		{
			flag = FormulaFlow.D(a);
			if (!flag)
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
				flag = FormulaFlow.B(a);
			}
			flag2 = FormulaFlow.C(a);
			if (!flag2)
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
				flag2 = FormulaFlow.A(a);
			}
		}
		else
		{
			try
			{
				if (A.Column == D)
				{
					flag = FormulaFlow.B(a, C);
				}
				else if (A.Column == E)
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
					flag = FormulaFlow.D(a, C);
				}
				else if (Conversions.ToBoolean(A.get_Offset((object)0, (object)(-1)).HasFormula))
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
					flag = ((!H.Where([SpecialName] (Range range) => Operators.CompareString(range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), A.get_Offset((object)0, (object)(-1)).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0).Any()) ? FormulaFlow.D(a, C) : FormulaFlow.B(a, C));
				}
				else
				{
					flag = FormulaFlow.B(a, C);
				}
				if (!flag)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						H.Add(A);
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				flag = false;
				ProjectData.ClearProjectError();
			}
			try
			{
				if (A.Row == F)
				{
					flag2 = FormulaFlow.A(a, C);
				}
				else if (A.Row == G)
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
					flag2 = FormulaFlow.C(a, C);
				}
				else if (Conversions.ToBoolean(A.get_Offset((object)(-1), (object)0).HasFormula))
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
					if (I.Where([SpecialName] (Range range) => Operators.CompareString(range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), A.get_Offset((object)(-1), (object)0).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0).Any())
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
						flag2 = FormulaFlow.A(a, C);
					}
					else
					{
						flag2 = FormulaFlow.C(a, C);
					}
				}
				else
				{
					flag2 = FormulaFlow.A(a, C);
				}
				if (!flag2)
				{
					I.Add(A);
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				flag2 = false;
				ProjectData.ClearProjectError();
			}
		}
		if (flag)
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
			if (flag2)
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
				FormulaFlow.G(A);
				goto IL_02f3;
			}
		}
		if (flag)
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
			FormulaFlow.F(A);
		}
		else if (flag2)
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
			FormulaFlow.E(A);
		}
		else
		{
			FormulaFlow.H(A);
		}
		goto IL_02f3;
		IL_02f3:
		a = null;
	}

	private static void C(Range A)
	{
		Interior interior = A.Interior;
		if (Operators.ConditionalCompareObjectEqual(interior.PatternColor, Common.PATTERN_COLOR_BLUE, TextCompare: false) || Operators.ConditionalCompareObjectEqual(interior.PatternColor, PATTERN_COLOR_VIOLET, TextCompare: false))
		{
			interior.PatternColor = 0;
			Common.B(A);
		}
		interior = null;
	}

	private static bool A(Range A, Worksheet B)
	{
		bool result = false;
		Range range = A.get_Offset((object)1, (object)0);
		string text = Conversions.ToString(((_Worksheet)B).get_Range((object)range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula);
		((_Worksheet)range.Worksheet).get_Range((object)A, (object)A.get_Offset((object)1, (object)0)).FillDown();
		if (Operators.ConditionalCompareObjectEqual(range.Formula, text, TextCompare: false))
		{
			result = true;
		}
		range.Formula = text;
		range = null;
		return result;
	}

	private static bool B(Range A, Worksheet B)
	{
		bool result = false;
		Range range = A.get_Offset((object)0, (object)1);
		string text = Conversions.ToString(((_Worksheet)B).get_Range((object)range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula);
		((_Worksheet)range.Worksheet).get_Range((object)A, (object)A.get_Offset((object)0, (object)1)).FillRight();
		if (Operators.ConditionalCompareObjectEqual(range.Formula, text, TextCompare: false))
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
			result = true;
		}
		range.Formula = text;
		range = null;
		return result;
	}

	private static bool C(Range A, Worksheet B)
	{
		bool result = false;
		Range range = A;
		string text = Conversions.ToString(((_Worksheet)B).get_Range((object)range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula);
		((_Worksheet)range.Worksheet).get_Range((object)range.get_Offset((object)(-1), (object)0), (object)A).FillDown();
		if (Operators.ConditionalCompareObjectEqual(range.Formula, text, TextCompare: false))
		{
			result = true;
		}
		range.Formula = text;
		range = null;
		return result;
	}

	private static bool D(Range A, Worksheet B)
	{
		bool result = false;
		Range range = A;
		string text = Conversions.ToString(((_Worksheet)B).get_Range((object)range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).Formula);
		((_Worksheet)range.Worksheet).get_Range((object)range.get_Offset((object)0, (object)(-1)), (object)A).FillRight();
		if (Operators.ConditionalCompareObjectEqual(range.Formula, text, TextCompare: false))
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
		range.Formula = text;
		range = null;
		return result;
	}

	private static bool A(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.Formula, A.get_Offset((object)1, (object)0).Formula, TextCompare: false);
	}

	private static bool B(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.Formula, A.get_Offset((object)0, (object)1).Formula, TextCompare: false);
	}

	private static bool C(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.Formula, A.get_Offset((object)(-1), (object)0).Formula, TextCompare: false);
	}

	private static bool D(Range A)
	{
		return Operators.ConditionalCompareObjectEqual(A.Formula, A.get_Offset((object)0, (object)(-1)).Formula, TextCompare: false);
	}

	private static void D(Range A)
	{
		A.Interior.Pattern = XlPattern.xlPatternLightDown;
		I(A);
	}

	private static void E(Range A)
	{
		A.Interior.Pattern = XlPattern.xlPatternLightVertical;
		I(A);
	}

	private static void F(Range A)
	{
		A.Interior.Pattern = XlPattern.xlPatternLightHorizontal;
		I(A);
	}

	private static void G(Range A)
	{
		A.Interior.Pattern = XlPattern.xlPatternGrid;
		I(A);
	}

	private static void H(Range A)
	{
		A.Interior.Pattern = XlPattern.xlPatternGray75;
		I(A);
	}

	private static void I(Range A)
	{
		A.Interior.PatternColor = Common.PATTERN_COLOR_BLUE;
	}
}
