using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using ExcelAddIn1.Audit;
using ExcelAddIn1.Format;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.Model;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.ExcelApp;

public sealed class Events
{
	public static void Add()
	{
		Application application = MH.A.Application;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(51213)).AddEventHandler(application, new AppEvents_SheetChangeEventHandler(B));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82507)).AddEventHandler(application, new AppEvents_ProtectedViewWindowActivateEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82562)).AddEventHandler(application, new AppEvents_ProtectedViewWindowDeactivateEventHandler(B));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82621)).AddEventHandler(application, new AppEvents_ProtectedViewWindowOpenEventHandler(C));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82668)).AddEventHandler(application, new AppEvents_ProtectedViewWindowBeforeCloseEventHandler(A));
		application = null;
	}

	public static void Remove()
	{
		Application application = MH.A.Application;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(application, new AppEvents_SheetSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(51213)).RemoveEventHandler(application, new AppEvents_SheetChangeEventHandler(B));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82507)).RemoveEventHandler(application, new AppEvents_ProtectedViewWindowActivateEventHandler(A));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82562)).RemoveEventHandler(application, new AppEvents_ProtectedViewWindowDeactivateEventHandler(B));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82621)).RemoveEventHandler(application, new AppEvents_ProtectedViewWindowOpenEventHandler(C));
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(82668)).RemoveEventHandler(application, new AppEvents_ProtectedViewWindowBeforeCloseEventHandler(A));
		application = null;
	}

	private static void A(object A, Range B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
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
				case 582:
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
							goto IL_001b;
						case 4:
							goto IL_002f;
						case 5:
							goto IL_0041;
						case 6:
							goto IL_0055;
						case 7:
							goto IL_0067;
						case 8:
							goto IL_007b;
						case 9:
							goto IL_008d;
						case 10:
							goto IL_00a2;
						case 11:
							goto IL_00b7;
						case 12:
							goto IL_00ca;
						case 13:
							goto IL_00df;
						case 14:
							goto IL_00e8;
						case 15:
							goto IL_00f1;
						case 16:
							goto IL_00fa;
						case 17:
							goto IL_0103;
						case 18:
							goto IL_010c;
						case 19:
							goto IL_0115;
						case 20:
							goto IL_011e;
						case 21:
							goto IL_0127;
						case 22:
							goto IL_0130;
						case 23:
							goto IL_0139;
						case 24:
							goto IL_0142;
						case 25:
							goto IL_014b;
						case 26:
							goto IL_0154;
						case 27:
							goto IL_015d;
						case 28:
							goto IL_0166;
						case 29:
							goto IL_016f;
						case 30:
							goto IL_0178;
						case 31:
							goto IL_0181;
						case 32:
							goto IL_018a;
						case 33:
							goto IL_0193;
						case 34:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 35:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0193:
					num2 = 33;
					AutoFill.CycleIndex = 0;
					break;
					IL_0007:
					num2 = 2;
					KH.A.CycleNumber.Index = 0;
					goto IL_001b;
					IL_001b:
					num2 = 3;
					KH.A.CycleCurrency.Index = 0;
					goto IL_002f;
					IL_002f:
					num2 = 4;
					KH.A.CyclePercent.Index = 0;
					goto IL_0041;
					IL_0041:
					num2 = 5;
					KH.A.CycleMultiple.Index = 0;
					goto IL_0055;
					IL_0055:
					num2 = 6;
					KH.A.CycleDate.Index = 0;
					goto IL_0067;
					IL_0067:
					num2 = 7;
					KH.A.CycleBinary.Index = 0;
					goto IL_007b;
					IL_007b:
					num2 = 8;
					KH.A.CycleRatio.Index = 0;
					goto IL_008d;
					IL_008d:
					num2 = 9;
					KH.A.FontColorCycle.Index = 0;
					goto IL_00a2;
					IL_00a2:
					num2 = 10;
					KH.A.FillColorCycle.Index = 0;
					goto IL_00b7;
					IL_00b7:
					num2 = 11;
					KH.A.BorderColorCycle.Index = 0;
					goto IL_00ca;
					IL_00ca:
					num2 = 12;
					KH.A.ChartColorCycle.Index = 0;
					goto IL_00df;
					IL_00df:
					num2 = 13;
					Alignment.HorizontalCycleIndex = 0;
					goto IL_00e8;
					IL_00e8:
					num2 = 14;
					Alignment.VerticalCycleIndex = 0;
					goto IL_00f1;
					IL_00f1:
					num2 = 15;
					FontColor.AutoColorIndex = 0;
					goto IL_00fa;
					IL_00fa:
					num2 = 16;
					FontStyle.CycleIndex = 0;
					goto IL_0103;
					IL_0103:
					num2 = 17;
					FontSize.CycleIndex = 0;
					goto IL_010c;
					IL_010c:
					num2 = 18;
					ExcelAddIn1.Format.Borders.BorderIndex = 0;
					goto IL_0115;
					IL_0115:
					num2 = 19;
					ExcelAddIn1.Format.Borders.CycleIndex = 0;
					goto IL_011e;
					IL_011e:
					num2 = 20;
					ExcelAddIn1.Format.Borders.OutsideBorderCycleIndex = 0;
					goto IL_0127;
					IL_0127:
					num2 = 21;
					ExcelAddIn1.Format.Borders.InsideBorderCycleIndex = 0;
					goto IL_0130;
					IL_0130:
					num2 = 22;
					ExcelAddIn1.Format.Styles.CycleNumber = 0;
					goto IL_0139;
					IL_0139:
					num2 = 23;
					ExcelAddIn1.Format.Styles.CycleIndex = 0;
					goto IL_0142;
					IL_0142:
					num2 = 24;
					CellSize.RowCycleIndex = 0;
					goto IL_014b;
					IL_014b:
					num2 = 25;
					CellSize.ColumnCycleIndex = 0;
					goto IL_0154;
					IL_0154:
					num2 = 26;
					Underline.CycleIndex = 0;
					goto IL_015d;
					IL_015d:
					num2 = 27;
					Footnotes.CycleIndex = 0;
					goto IL_0166;
					IL_0166:
					num2 = 28;
					Lists.CycleIndex = 0;
					goto IL_016f;
					IL_016f:
					num2 = 29;
					AlternateShading.CycleIndex = 0;
					goto IL_0178;
					IL_0178:
					num2 = 30;
					Paintbrush.CycleIndex = 0;
					goto IL_0181;
					IL_0181:
					num2 = 31;
					TraceAll.ShowingPrecedents = false;
					goto IL_018a;
					IL_018a:
					num2 = 32;
					TraceAll.ShowingDependents = false;
					goto IL_0193;
					end_IL_0000_2:
					break;
				}
				num2 = 34;
				FastFill.CycleIndex = 0;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 582;
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
	}

	private static void B(object A, Range B)
	{
		if (!KH.A.AutoColorOnEntry || !(A is Worksheet))
		{
			return;
		}
		bool flag = default(bool);
		bool flag2 = default(bool);
		bool flag4 = default(bool);
		IEnumerator enumerator = default(IEnumerator);
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
			Application application = B.Application;
			if (application.CutCopyMode == XlCutCopyMode.xlCut)
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
			if (Operators.CompareString(((Microsoft.Office.Interop.Excel.Workbook)NewLateBinding.LateGet(A, null, VH.A(8701), new object[0], null, null, null)).Name, application.ActiveWorkbook.Name, TextCompare: false) != 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						application = null;
						return;
					}
				}
			}
			Range range = B;
			if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(range.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51236), new object[0], null, null, null), null, VH.A(5814), new object[0], null, null, null), range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
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
				flag = true;
			}
			if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51255), new object[0], null, null, null), null, VH.A(5814), new object[0], null, null, null), range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
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
				flag2 = true;
			}
			range = null;
			if (!flag)
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
				if (!flag2)
				{
					Range range2 = B;
					if (Operators.ConditionalCompareObjectEqual(range2.EntireRow.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), NewLateBinding.LateGet(application.Selection, null, VH.A(5814), new object[0], null, null, null), TextCompare: false))
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
					if (Operators.ConditionalCompareObjectEqual(range2.EntireColumn.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), NewLateBinding.LateGet(application.Selection, null, VH.A(5814), new object[0], null, null, null), TextCompare: false))
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
					range2 = null;
					string text;
					try
					{
						text = Conversions.ToString(NewLateBinding.LateGet(application.CommandBars[VH.A(82729)].Controls[VH.A(82746)].Control, null, VH.A(82757), new object[1] { 1 }, null, null, null));
						uint num = TH.A(text);
						if (num <= 790754359)
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
							if (num <= 96404597)
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
								if (num != 72689141)
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
									if (num == 96404597)
									{
										if (Operators.CompareString(text, VH.A(82766), TextCompare: false) == 0)
										{
											goto IL_054e;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_0472;
											}
											continue;
											end_IL_0472:
											break;
										}
									}
								}
								else if (Operators.CompareString(text, VH.A(82871), TextCompare: false) == 0)
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
									goto IL_054e;
								}
							}
							else if (num != 135637716)
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
								if (num == 790754359 && Operators.CompareString(text, VH.A(82791), TextCompare: false) == 0)
								{
									goto IL_054e;
								}
							}
							else
							{
								if (Operators.CompareString(text, VH.A(82856), TextCompare: false) == 0)
								{
									goto IL_054e;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_0523;
									}
									continue;
									end_IL_0523:
									break;
								}
							}
						}
						else if (num <= 1589776164)
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
							if (num != 1469573738)
							{
								if (num != 1589776164)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_03f6;
										}
										continue;
										end_IL_03f6:
										break;
									}
								}
								else
								{
									if (Operators.CompareString(text, VH.A(82808), TextCompare: false) == 0)
									{
										goto IL_054e;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_04df;
										}
										continue;
										end_IL_04df:
										break;
									}
								}
							}
							else
							{
								if (Operators.CompareString(text, VH.A(60691), TextCompare: false) == 0)
								{
									goto IL_054e;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_049b;
									}
									continue;
									end_IL_049b:
									break;
								}
							}
						}
						else if (num != 3156332295u)
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
							if (num == 3839184739u)
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
								if (Operators.CompareString(text, VH.A(65312), TextCompare: false) == 0)
								{
									goto IL_054e;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0449;
									}
									continue;
									end_IL_0449:
									break;
								}
							}
						}
						else
						{
							if (Operators.CompareString(text, VH.A(82833), TextCompare: false) == 0)
							{
								goto IL_054e;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0502;
								}
								continue;
								end_IL_0502:
								break;
							}
						}
						goto end_IL_02e9;
						IL_054e:
						application = null;
						return;
						end_IL_02e9:;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						text = VH.A(82916);
						ProjectData.ClearProjectError();
					}
					Application application2 = application;
					Range activeCell = application2.ActiveCell;
					application2.EnableCancelKey = XlEnableCancelKey.xlDisabled;
					try
					{
						if (KH.A.UndoEnabled)
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
							application2.ScreenUpdating = false;
							application2.EnableEvents = false;
							bool flag3 = application2.CutCopyMode == XlCutCopyMode.xlCopy;
							application2.DisplayAlerts = false;
							XlCalculation calculation = application2.Calculation;
							application2.Calculation = XlCalculation.xlCalculationManual;
							try
							{
								application2.Undo();
								if (application2.CutCopyMode != XlCutCopyMode.xlCut)
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
									flag4 = JH.A(B);
								}
								application2.Undo();
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							application2.DisplayAlerts = true;
							application2.Calculation = calculation;
							try
							{
								if (flag3)
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										Paste.CopiedRange.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
										break;
									}
								}
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
							}
						}
						if (application2.CutCopyMode != XlCutCopyMode.xlCopy)
						{
							if (application2.ScreenUpdating)
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
								application2.ScreenUpdating = false;
							}
							try
							{
								enumerator = B.GetEnumerator();
								try
								{
									while (enumerator.MoveNext())
									{
										AutoColor.AutoColorIfNotEmpty((Range)enumerator.Current);
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_06c0;
										}
										continue;
										end_IL_06c0:
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
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
							}
							application2.EnableEvents = false;
						}
						try
						{
							activeCell.Activate();
						}
						catch (Exception ex9)
						{
							ProjectData.SetProjectError(ex9);
							Exception ex10 = ex9;
							ProjectData.ClearProjectError();
						}
						if (flag4)
						{
							JH.A(B, text);
						}
					}
					catch (Exception ex11)
					{
						ProjectData.SetProjectError(ex11);
						Exception ex12 = ex11;
						ProjectData.ClearProjectError();
					}
					application2.EnableCancelKey = XlEnableCancelKey.xlInterrupt;
					application2.EnableEvents = true;
					application2.ScreenUpdating = true;
					application2 = null;
					activeCell = null;
					application = null;
					return;
				}
			}
			application = null;
			return;
		}
	}

	private static void A(ProtectedViewWindow A)
	{
		KH.A.Invalidate();
	}

	private static void B(ProtectedViewWindow A)
	{
		KH.A.Invalidate();
	}

	private static void C(ProtectedViewWindow A)
	{
		KH.A.Invalidate();
	}

	private static void A(ProtectedViewWindow A, XlProtectedViewCloseReason B, ref bool C)
	{
		KH.A.Invalidate();
	}
}
