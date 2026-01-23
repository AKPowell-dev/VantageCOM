using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class SummaryStats
{
	public static void Add()
	{
		if (!Access.AllowExcelOperation((PlanType)5, (Restriction)2, false))
		{
			return;
		}
		checked
		{
			IEnumerator<Statistic> enumerator = default(IEnumerator<Statistic>);
			Range range2 = default(Range);
			string text = default(string);
			IEnumerator enumerator3 = default(IEnumerator);
			Range range5 = default(Range);
			IEnumerator enumerator4 = default(IEnumerator);
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
				Application application = MH.A.Application;
				Common.ClearVisualizations(application);
				if (application.Selection is Range)
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
					wpfSummaryStats wpfSummaryStats2 = new wpfSummaryStats();
					wpfSummaryStats2.ShowDialog();
					if (wpfSummaryStats2.DialogResult.HasValue)
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
						if (wpfSummaryStats2.DialogResult.Value)
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
							List<string> list = new List<string>();
							List<string> list2 = new List<string>();
							try
							{
								enumerator = wpfSummaryStats2.SummaryStats.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Statistic current = enumerator.Current;
									if (!current.IsChecked)
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
									list2.Add(current.Label);
									list.Add(current.FunctionEnglish);
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_0111;
									}
									continue;
									end_IL_0111:
									break;
								}
							}
							finally
							{
								if (enumerator != null)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										enumerator.Dispose();
										break;
									}
								}
							}
							bool value = wpfSummaryStats2.chkIgnoreCurrency.IsChecked.Value;
							Range range = (Range)application.Selection;
							int count = list.Count;
							XlCalculation calculation = application.Calculation;
							application.ScreenUpdating = false;
							application.EnableEvents = false;
							application.Calculation = XlCalculation.xlCalculationManual;
							try
							{
								((_Worksheet)range.Worksheet).get_Range(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range.Rows[RuntimeHelpers.GetObjectValue(range.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[2] { 1, 0 }, null, null, null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range.Rows[RuntimeHelpers.GetObjectValue(range.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[2]
								{
									count + 2,
									0
								}, null, null, null))).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
								range2 = ((Range)range.Rows[RuntimeHelpers.GetObjectValue(range.Rows.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Offset((object)2, (object)0).get_Resize((object)count, RuntimeHelpers.GetObjectValue(Missing.Value));
								int num = 1;
								foreach (Range cell in ((Range)range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Cells)
								{
									if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(cell.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
										if (!Information.IsDate(RuntimeHelpers.GetObjectValue(cell.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
											if (!value)
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
											try
											{
												text = cell.Text.ToString();
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												ProjectData.ClearProjectError();
											}
											if (!text.Contains(VH.A(41262)))
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
												if (!text.Contains(VH.A(96009)) && !text.Contains(VH.A(96012)))
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
													if (!text.Contains(VH.A(96015)))
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
														break;
													}
												}
											}
										}
									}
									num++;
								}
								int num2;
								if (num == 1)
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
									if (range.Column > 1)
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
										range2 = range2.get_Offset((object)0, (object)(-1)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.AddObject(range2.Columns.CountLarge, 1));
										num = 2;
										num2 = -1;
										goto IL_04d6;
									}
								}
								num2 = 0;
								goto IL_04d6;
								IL_04d6:
								Range range4 = JH.A(range2, application);
								if (range4 != null)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										int row = ((Range)range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Row;
										int row2 = ((Range)range.Rows[2, RuntimeHelpers.GetObjectValue(Missing.Value)]).Row;
										int num3 = 0;
										try
										{
											enumerator3 = range4.Rows.GetEnumerator();
											while (enumerator3.MoveNext())
											{
												Range obj = (Range)enumerator3.Current;
												range5 = obj.get_Offset((object)0, (object)(num - 1)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.AddObject(Operators.SubtractObject(range2.Columns.CountLarge, num), 1));
												NewLateBinding.LateSetComplex(obj.Cells[1, 1], null, VH.A(57302), new object[1] { list2[num3] }, null, null, OptimisticSet: false, RValueBase: true);
												string text2 = ((Range)range.Columns[num + num2, RuntimeHelpers.GetObjectValue(Missing.Value)]).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
												string text3 = list[num3];
												string text4;
												if (text3.Contains(VH.A(96018)))
												{
													text4 = list2[num3].ToLower();
													if (!text4.Contains(VH.A(96037)))
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
														if (!text4.Contains(VH.A(96040)))
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
															if (!text4.Contains(VH.A(96051)))
															{
																if (!text4.Contains(VH.A(96073)))
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
																	if (!text4.Contains(VH.A(96076)))
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
																		if (!text4.Contains(VH.A(96087)))
																		{
																			range5.Formula = VH.A(75261);
																			goto IL_07f1;
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
																}
																range5.Formula = VH.A(48936) + text3 + VH.A(39848) + text2 + VH.A(96100);
																goto IL_07f1;
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
													range5.Formula = VH.A(48936) + text3 + VH.A(39848) + text2 + VH.A(96064);
													goto IL_07f1;
												}
												range5.Formula = VH.A(48936) + text3 + VH.A(39848) + text2 + VH.A(39904);
												goto IL_083e;
												IL_083e:
												int num4;
												if (num3 != 0)
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
													num4 = row2;
												}
												else
												{
													num4 = row;
												}
												int num5 = num4 - range5.Row;
												try
												{
													enumerator4 = range5.Cells.GetEnumerator();
													while (enumerator4.MoveNext())
													{
														Range obj2 = (Range)enumerator4.Current;
														obj2.NumberFormat = RuntimeHelpers.GetObjectValue(obj2.get_Offset((object)num5, RuntimeHelpers.GetObjectValue(Missing.Value)).NumberFormat);
													}
													while (true)
													{
														switch (3)
														{
														case 0:
															break;
														default:
															goto end_IL_08b1;
														}
														continue;
														end_IL_08b1:
														break;
													}
												}
												finally
												{
													if (enumerator4 is IDisposable)
													{
														while (true)
														{
															switch (3)
															{
															case 0:
																continue;
															}
															(enumerator4 as IDisposable).Dispose();
															break;
														}
													}
												}
												num3++;
												continue;
												IL_07f1:
												text4 = null;
												goto IL_083e;
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													break;
												default:
													goto end_IL_08ef;
												}
												continue;
												end_IL_08ef:
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
										try
										{
											AutoColor.SetColor(range2.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value)), 2);
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											ProjectData.ClearProjectError();
										}
										long num6 = Conversions.ToLong(Operators.SubtractObject(NewLateBinding.LateGet(range2.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null), NewLateBinding.LateGet(range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null)));
										enumerator5 = range.Columns.GetEnumerator();
										try
										{
											while (enumerator5.MoveNext())
											{
												Range range6 = (Range)enumerator5.Current;
												if (application.WorksheetFunction.CountA(range6, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == 0.0)
												{
													range6.get_Offset((object)num6, RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize((object)count, RuntimeHelpers.GetObjectValue(Missing.Value)).Clear();
												}
											}
											while (true)
											{
												switch (4)
												{
												case 0:
													break;
												default:
													goto end_IL_0ba0;
												}
												continue;
												end_IL_0ba0:
												break;
											}
										}
										finally
										{
											IDisposable disposable = enumerator5 as IDisposable;
											if (disposable != null)
											{
												disposable.Dispose();
											}
										}
										range4 = null;
										break;
									}
								}
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								Forms.ErrorMessage(VH.A(96109));
								ProjectData.ClearProjectError();
							}
							application.Calculation = calculation;
							application.ScreenUpdating = true;
							application.EnableEvents = true;
							JH.A((object)range);
							JH.A((object)range2);
							JH.A((object)range5);
							list = null;
							list2 = null;
							clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(96228));
						}
					}
					wpfSummaryStats2.SummaryStats = null;
					wpfSummaryStats2 = null;
				}
				application = null;
				return;
			}
		}
	}
}
