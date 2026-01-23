using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Audit.Visualizations;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class Scenarios
{
	private static readonly string m_A = VH.A(95992);

	[CompilerGenerated]
	private static ObservableCollection<string> m_A;

	[CompilerGenerated]
	private static bool m_A;

	private static ObservableCollection<string> LastUsedLocalScenarios
	{
		[CompilerGenerated]
		get
		{
			return Scenarios.m_A;
		}
		[CompilerGenerated]
		set
		{
			Scenarios.m_A = value;
		}
	} = null;

	private static bool UseLocalScenariosByDefault
	{
		[CompilerGenerated]
		get
		{
			return Scenarios.m_A;
		}
		[CompilerGenerated]
		set
		{
			Scenarios.m_A = value;
		}
	} = false;

	public static void Add()
	{
		if (!Access.AllowExcelOperation((PlanType)5, (Restriction)2, false))
		{
			return;
		}
		checked
		{
			Range range = default(Range);
			Range range2 = default(Range);
			Range range3 = default(Range);
			Range range4 = default(Range);
			IEnumerator enumerator = default(IEnumerator);
			Range range6 = default(Range);
			Range range8 = default(Range);
			Range a = default(Range);
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
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				Common.ClearVisualizations(application);
				if (application.Selection is Range)
				{
					if (application.ActiveWindow.SelectedSheets.Count == 1)
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
						range = (Range)application.Selection;
						ExcelAddIn1.Sheets.Protection.Unprotect(range.Worksheet);
						if (!range.Worksheet.ProtectContents)
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
							if (Operators.ConditionalCompareObjectEqual(range.Rows.CountLarge, 1, TextCompare: false))
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
									if (((Range)range.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Column > 3)
									{
										range2 = null;
										range3 = A(application.ActiveWorkbook);
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
												range2 = ((_Worksheet)range3.Worksheet).get_Range((object)range3.Validation.Formula1.Substring(1), RuntimeHelpers.GetObjectValue(Missing.Value));
											}
											catch (Exception ex)
											{
												ProjectData.SetProjectError(ex);
												Exception ex2 = ex;
												ProjectData.ClearProjectError();
											}
										}
										if (LastUsedLocalScenarios == null)
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
											LastUsedLocalScenarios = new ObservableCollection<string>();
											ObservableCollection<string> lastUsedLocalScenarios = LastUsedLocalScenarios;
											lastUsedLocalScenarios.Add(VH.A(95171));
											lastUsedLocalScenarios.Add(VH.A(95194));
											lastUsedLocalScenarios.Add(VH.A(95225));
											_ = null;
										}
										wpfScenarios wpfScenarios2 = new wpfScenarios();
										wpfScenarios2.GlobalScenariosFound = range2 != null;
										if (UseLocalScenariosByDefault)
										{
											wpfScenarios2.radLocal.IsChecked = true;
										}
										else
										{
											wpfScenarios2.radGlobal.IsChecked = true;
										}
										ObservableCollection<string> observableCollection = new ObservableCollection<string>();
										if (range2 != null)
										{
											wpfScenarios2.radGlobal.Content = VH.A(95252);
											if (Operators.ConditionalCompareObjectEqual(range2.Columns.CountLarge, 1, TextCompare: false))
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
												range4 = range2.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)1);
											}
											else
											{
												range4 = range2.get_Offset((object)1, RuntimeHelpers.GetObjectValue(Missing.Value));
											}
											try
											{
												enumerator = range4.GetEnumerator();
												while (enumerator.MoveNext())
												{
													Range range5 = (Range)enumerator.Current;
													try
													{
														observableCollection.Add(range5.Text.ToString());
													}
													catch (Exception ex3)
													{
														ProjectData.SetProjectError(ex3);
														Exception ex4 = ex3;
														observableCollection.Add("");
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
														goto end_IL_02f9;
													}
													continue;
													end_IL_02f9:
													break;
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
														(enumerator as IDisposable).Dispose();
														break;
													}
												}
											}
										}
										else
										{
											wpfScenarios2.radGlobal.Content = VH.A(95311);
											observableCollection.Add(VH.A(95171));
											observableCollection.Add(VH.A(95194));
											observableCollection.Add(VH.A(95225));
										}
										wpfScenarios2.lvGlobal.ItemsSource = observableCollection;
										observableCollection = null;
										wpfScenarios2.lvLocal.ItemsSource = LastUsedLocalScenarios;
										wpfScenarios2.ShowDialog();
										if (wpfScenarios2.DialogResult.HasValue)
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
											if (wpfScenarios2.DialogResult.Value)
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
												if (!application.ActiveWorkbook.Saved)
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
													if (MessageBox.Show(VH.A(95366), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
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
														application.ActiveWorkbook.Save();
													}
												}
												int count;
												if (wpfScenarios2.radGlobal.IsChecked == true)
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
													UseLocalScenariosByDefault = false;
													count = wpfScenarios2.lvGlobal.Items.Count;
												}
												else
												{
													UseLocalScenariosByDefault = true;
													count = wpfScenarios2.lvLocal.Items.Count;
												}
												application.ScreenUpdating = false;
												application.EnableEvents = false;
												XlCalculation calculation = application.Calculation;
												application.Calculation = XlCalculation.xlCalculationManual;
												try
												{
													range.get_Offset((object)1, RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize((object)(count + 2), RuntimeHelpers.GetObjectValue(Missing.Value)).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
													range4 = range.get_Offset((object)2, (object)(-2)).get_Resize((object)count, (object)1);
													string text2;
													int num2;
													if (wpfScenarios2.radGlobal.IsChecked == true)
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
														if (range2 != null)
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
															bool flag;
															if (Operators.ConditionalCompareObjectEqual(range2.Columns.CountLarge, 1, TextCompare: false))
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
																flag = true;
															}
															else
															{
																flag = false;
															}
															int num = count;
															for (int i = 1; i <= num; i++)
															{
																object instance = range2.Cells[i, RuntimeHelpers.GetObjectValue(Missing.Value)];
																string text;
																if (flag)
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
																	object instance2 = NewLateBinding.LateGet(instance, null, VH.A(60565), new object[2]
																	{
																		Missing.Value,
																		1
																	}, null, null, null);
																	text = ((range4.Worksheet != range2.Worksheet) ? Conversions.ToString(Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(instance2, null, VH.A(5814), new object[4]
																	{
																		1,
																		1,
																		XlReferenceStyle.xlA1,
																		true
																	}, null, null, null))) : Conversions.ToString(Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(instance2, null, VH.A(5814), new object[2] { 1, 1 }, null, null, null))));
																	instance2 = null;
																}
																else
																{
																	object instance3 = NewLateBinding.LateGet(instance, null, VH.A(60565), new object[2]
																	{
																		1,
																		Missing.Value
																	}, null, null, null);
																	if (range4.Worksheet == range2.Worksheet)
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
																		text = Conversions.ToString(Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(instance3, null, VH.A(5814), new object[2] { 1, 1 }, null, null, null)));
																	}
																	else
																	{
																		text = Conversions.ToString(Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(instance3, null, VH.A(5814), new object[4]
																		{
																			1,
																			1,
																			XlReferenceStyle.xlA1,
																			true
																		}, null, null, null)));
																	}
																	instance3 = null;
																}
																instance = null;
																NewLateBinding.LateSetComplex(range4.Cells[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { text }, null, null, OptimisticSet: false, RValueBase: true);
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
															text2 = Scenarios.m_A;
															NewLateBinding.LateSetComplex(NewLateBinding.LateGet(range4.Cells[RuntimeHelpers.GetObjectValue(range4.Cells.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[2]
															{
																1,
																Missing.Value
															}, null, null, null), null, VH.A(68956), new object[1] { VH.A(95526) + range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + text2 + VH.A(39904) }, null, null, OptimisticSet: false, RValueBase: true);
															num2 = 2;
														}
														else
														{
															range2 = range.get_Offset((object)2, (object)(-3)).get_Resize((object)count, (object)1);
															range3 = (Range)NewLateBinding.LateGet(range2.Cells[RuntimeHelpers.GetObjectValue(range2.Cells.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[2]
															{
																1,
																Missing.Value
															}, null, null, null);
															try
															{
																application.ActiveWorkbook.Names.Add(Scenarios.m_A, range3, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
															}
															catch (Exception ex5)
															{
																ProjectData.SetProjectError(ex5);
																Exception ex6 = ex5;
																Forms.ErrorMessage(ex6.Message);
																throw ex6;
															}
															text2 = Scenarios.m_A;
															A(wpfScenarios2.lvGlobal, range2, range4, range3, text2);
															num2 = 3;
														}
													}
													else
													{
														range2 = range.get_Offset((object)2, (object)(-3)).get_Resize((object)count, (object)1);
														range3 = (Range)NewLateBinding.LateGet(range2.Cells[RuntimeHelpers.GetObjectValue(range2.Cells.CountLarge), RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[2]
														{
															1,
															Missing.Value
														}, null, null, null);
														text2 = range3.get_Address((object)0, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
														A(wpfScenarios2.lvLocal, range2, range4, range3, text2);
														num2 = 3;
													}
													range6 = range.get_Offset((object)2, RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize((object)count, RuntimeHelpers.GetObjectValue(Missing.Value));
													range6.Value2 = 0;
													A(range6);
													int num3 = Conversions.ToInteger(range3.Value2);
													bool flag2 = true;
													try
													{
														Range range7 = null;
														try
														{
															range7 = range.SpecialCells(XlCellType.xlCellTypeBlanks, RuntimeHelpers.GetObjectValue(Missing.Value));
														}
														catch (Exception ex7)
														{
															ProjectData.SetProjectError(ex7);
															Exception ex8 = ex7;
															ProjectData.ClearProjectError();
														}
														if (range7 != null)
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
															if (!Operators.ConditionalCompareObjectLess(range7.CountLarge, range.Cells.CountLarge, TextCompare: false))
															{
																flag2 = false;
																goto IL_0c38;
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
														NewLateBinding.LateSetComplex(range6.Rows[num3, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { range.Formula }, null, null, OptimisticSet: false, RValueBase: true);
														goto IL_0c38;
														IL_0c38:
														range7 = null;
													}
													catch (Exception ex9)
													{
														ProjectData.SetProjectError(ex9);
														Exception ex10 = ex9;
														flag2 = false;
														ProjectData.ClearProjectError();
													}
													range8 = range.get_Offset((object)(count + 2), RuntimeHelpers.GetObjectValue(Missing.Value));
													int num4 = Conversions.ToInteger(range.Columns.CountLarge);
													for (int j = 1; j <= num4; j++)
													{
														if (flag2)
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
																NewLateBinding.LateSetComplex(NewLateBinding.LateGet(range6.Cells[num3, j], null, VH.A(60635), new object[0], null, null, null), null, VH.A(55331), new object[1] { NewLateBinding.LateGet(NewLateBinding.LateGet(range.Cells[j, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60635), new object[0], null, null, null), null, VH.A(55331), new object[0], null, null, null) }, null, null, OptimisticSet: false, RValueBase: true);
															}
															catch (Exception ex11)
															{
																ProjectData.SetProjectError(ex11);
																Exception ex12 = ex11;
																ProjectData.ClearProjectError();
															}
														}
														NewLateBinding.LateSetComplex(range8.Cells[j, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(95526), NewLateBinding.LateGet(range6.Columns[j, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)), VH.A(2378)), text2), VH.A(39904)) }, null, null, OptimisticSet: false, RValueBase: true);
														NewLateBinding.LateSetComplex(range.Cells[j, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(68956), new object[1] { Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(range8.Cells[j, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null)) }, null, null, OptimisticSet: false, RValueBase: true);
													}
													range8.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
													range.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
													string numberFormat = (string)(range6.NumberFormat = Conversions.ToString(NewLateBinding.LateGet(range.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57240), new object[0], null, null, null)));
													range8.NumberFormat = numberFormat;
													range.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
													try
													{
														((Range)range6.Rows[num3, RuntimeHelpers.GetObjectValue(Missing.Value)]).PasteSpecial(XlPasteType.xlPasteComments, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
														range.ClearComments();
													}
													catch (Exception ex13)
													{
														ProjectData.SetProjectError(ex13);
														Exception ex14 = ex13;
														ProjectData.ClearProjectError();
													}
													application.CutCopyMode = (XlCutCopyMode)0;
													range.Select();
													Microsoft.Office.Interop.Excel.Border border = range8.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(-num2)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), Operators.AddObject(range8.Columns.CountLarge, num2)).Borders[XlBordersIndex.xlEdgeTop];
													border.Color = KH.A.DefaultBorderColor;
													border.LineStyle = XlLineStyle.xlDot;
													border.Weight = 1;
													_ = null;
												}
												catch (Exception ex15)
												{
													ProjectData.SetProjectError(ex15);
													Exception ex16 = ex15;
													Forms.ErrorMessage(ex16.Message);
													clsReporting.LogException(ex16);
													ProjectData.ClearProjectError();
												}
												application.ScreenUpdating = true;
												application.EnableEvents = true;
												application.Calculation = calculation;
												clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, VH.A(95541));
											}
										}
										wpfScenarios2 = null;
									}
									else
									{
										Forms.WarningMessage(VH.A(95568));
									}
									goto IL_111d;
								}
							}
							Forms.WarningMessage(VH.A(95663));
						}
						else
						{
							Forms.WarningMessage(VH.A(95776));
						}
					}
					else
					{
						Forms.WarningMessage(VH.A(95887));
					}
				}
				goto IL_111d;
				IL_111d:
				application = null;
				JH.A((object)range);
				JH.A((object)range3);
				JH.A((object)range2);
				JH.A((object)range4);
				JH.A((object)a);
				JH.A((object)range6);
				JH.A((object)range8);
				return;
			}
		}
	}

	private static Range A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		Range range = null;
		Names names = A.Names;
		Name name;
		while (range == null)
		{
			name = null;
			try
			{
				name = names.Item(Scenarios.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (name == null)
			{
				break;
			}
			if (name.Parent is Microsoft.Office.Interop.Excel.Workbook)
			{
				try
				{
					range = name.RefersToRange;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				if (range != null)
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
				break;
			}
			name.Delete();
			names = A.Names;
		}
		names = null;
		name = null;
		return range;
	}

	private static void A(System.Windows.Controls.ListBox A, Range B, Range C, Range D, string E)
	{
		int num = 1;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ((IEnumerable)A.Items).GetEnumerator();
			while (enumerator.MoveNext())
			{
				string text = Conversions.ToString(enumerator.Current);
				NewLateBinding.LateSetComplex(B.Cells[num, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[1] { num }, null, null, OptimisticSet: false, RValueBase: true);
				NewLateBinding.LateSetComplex(C.Cells[num, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57302), new object[1] { text }, null, null, OptimisticSet: false, RValueBase: true);
				num = checked(num + 1);
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
		D.Value2 = 1;
		D.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)1).Formula = VH.A(95526) + C.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(2378) + E + VH.A(39904);
		Microsoft.Office.Interop.Excel.Validation validation = D.Validation;
		validation.Delete();
		validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, VH.A(48936) + B.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value));
		validation.IgnoreBlank = false;
		validation.InCellDropdown = true;
		validation.ShowError = true;
		validation.ShowInput = true;
		_ = null;
		try
		{
			Scenarios.A(B);
			Scenarios.A(D);
			B.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			D.HorizontalAlignment = XlHAlign.xlHAlignCenter;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Range A)
	{
		if (KH.A.AutoColors[0].Length <= 0)
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
			A.Font.Color = clsColors.RGB2Ole(KH.A.AutoColors[0]);
			return;
		}
	}
}
