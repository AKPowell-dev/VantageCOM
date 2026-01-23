using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class StackedWaterfall
{
	private enum JD
	{
		A = 1,
		B,
		C
	}

	private enum KD
	{
		A = 1,
		B
	}

	private struct LD
	{
		public float A;

		public float B;

		public bool A;

		public string A;

		public bool B;

		public JD A;

		public int A;
	}

	private struct MD
	{
		public Range A;

		public Range B;

		public Range C;

		public Range D;

		public Range E;
	}

	private struct ND
	{
		public double A;

		public double B;

		public double C;

		public double D;
	}

	private static readonly string m_A = VH.A(62560);

	private static readonly int m_A = 15;

	private const long m_A = 10921638L;

	private const long m_B = 9671571L;

	private const long C = 14277081L;

	private const string m_B = "T2";

	public static void Create()
	{
		if (!Licensing.AllowQuickChartOperation())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		bool B = true;
		string text = VH.A(57237);
		checked
		{
			if (application.Selection is Range)
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
				if (!Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
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
					Range A = (Range)application.Selection;
					LD lD = StackedWaterfall.A(ref A, ref B);
					if (B)
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
						string text2 = string.Empty;
						bool B2 = default(bool);
						bool C = default(bool);
						if (Operators.CompareString(Versioned.TypeName(A), VH.A(41315), TextCompare: false) == 0)
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
							text2 = A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							if (!StackedWaterfall.A(text2, ref B2, ref C))
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
								B = false;
								text2 = null;
							}
						}
						if (B)
						{
							int f = QuickCharts2.InputColor();
							int g = QuickCharts2.LinkColor();
							XlCalculation calc = default(XlCalculation);
							QuickCharts2.PrepareExcel(application, ref calc);
							Worksheet worksheet;
							ChartObject chartObject;
							Chart chart;
							MD mD;
							try
							{
								Range range = ((_Application)application).get_Range((object)text2, RuntimeHelpers.GetObjectValue(Missing.Value));
								long num = 0L;
								if (B2)
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
									num = 1L;
								}
								long num2 = 0L;
								if (C)
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
									num2 = 1L;
								}
								int num3 = range.Rows.Count;
								int count = range.Columns.Count;
								string text3 = VH.A(20593);
								long num4 = num3;
								for (long num5 = 1L; num5 <= num4; num5++)
								{
									text3 = Conversions.ToString(NewLateBinding.LateGet(range.Cells[num5 + num, 2], null, VH.A(57240), new object[0], null, null, null));
									if (Operators.CompareString(text3, VH.A(20593), TextCompare: false) != 0)
									{
										break;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_0218;
										}
										continue;
										end_IL_0218:
										break;
									}
								}
								long num6 = count - 1 - num2;
								worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(application.ActiveSheet), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								chartObject = QuickCharts2.AddChart(worksheet, lD.A, lD.B);
								chartObject.Placement = XlPlacement.xlFreeFloating;
								int val = chartObject.BottomRightCell.Row + 2;
								text += Conversions.ToString(Math.Max(val, StackedWaterfall.m_A + 3));
								chart = chartObject.Chart;
								chart.ChartType = XlChartType.xlColumnStacked;
								mD = StackedWaterfall.A(worksheet, chartObject, lD);
								Range D = default(Range);
								StackedWaterfall.A(lD, worksheet, ref A, ref D, Math.Max(val, StackedWaterfall.m_A + 3), f, g);
								Range range2 = ((_Worksheet)worksheet).get_Range((object)text, RuntimeHelpers.GetObjectValue(Missing.Value));
								if (!B2)
								{
									range2.EntireRow.Insert(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									num3++;
									range2 = ((_Worksheet)worksheet).get_Range((object)text, RuntimeHelpers.GetObjectValue(Missing.Value));
									long num7 = num6;
									for (long num8 = 1L; num8 <= num7; num8++)
									{
										string value = Conversions.ToString(Strings.Chr((int)(64 + num8)));
										range2.get_Offset((object)0, (object)num8).Value2 = value;
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
								if (Strings.Len(RuntimeHelpers.GetObjectValue(range2.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(1 + num6)).Value2)) == 0)
								{
									range2.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(1 + num6)).Value2 = VH.A(57265);
								}
								if (application.WorksheetFunction.Count(range2.get_Offset((object)(num3 - 1), (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num6), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) > 0.0)
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
									range2.get_Offset((object)num3, (object)0).Value2 = VH.A(57276);
								}
								range2 = range2.get_Resize((object)num3, (object)count);
								Range currentRegion = range2.CurrentRegion;
								num3 = currentRegion.Rows.Count;
								long num9 = num3 - 1;
								for (int i = num3; i >= 3; i += -1)
								{
									NewLateBinding.LateCall(NewLateBinding.LateGet(currentRegion.Rows[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51236), new object[0], null, null, null), null, VH.A(57289), new object[0], null, null, null, IgnoreReturn: true);
									NewLateBinding.LateCall(NewLateBinding.LateGet(currentRegion.Rows[i, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51236), new object[0], null, null, null), null, VH.A(57289), new object[0], null, null, null, IgnoreReturn: true);
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									NewLateBinding.LateCall(NewLateBinding.LateGet(currentRegion.Rows[2, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(51236), new object[0], null, null, null), null, VH.A(57289), new object[0], null, null, null, IgnoreReturn: true);
									currentRegion = range2.get_Resize((object)(3 * num3 - 1), (object)(19 + 9 * num6));
									NewLateBinding.LateSetComplex(currentRegion.Cells[1, 3 + num6], null, VH.A(57302), new object[1] { VH.A(57315) }, null, null, OptimisticSet: false, RValueBase: true);
									object[,] array = new object[1, (int)(num6 - 1) + 1];
									object[,] array2 = new object[1, (int)(num6 - 1) + 1];
									int num10 = (int)(num6 - 1);
									for (int j = 0; j <= num10; j++)
									{
										array[0, j] = VH.A(57326) + Conversions.ToString(j + 2);
										array2[0, (int)(num6 - j - 1)] = VH.A(57326) + Conversions.ToString(j + 2);
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										object instance = currentRegion.Cells[1, 4 + num6];
										string memberName = VH.A(57333);
										object[] obj = new object[2]
										{
											Missing.Value,
											num6
										};
										object[] array3 = obj;
										bool[] obj2 = new bool[2] { false, true };
										bool[] array4 = obj2;
										object instance2 = NewLateBinding.LateGet(instance, null, memberName, obj, null, null, obj2);
										if (array4[1])
										{
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance2, null, VH.A(57346), new object[1] { array }, null, null, OptimisticSet: false, RValueBase: true);
										object[] array5 = new object[16]
										{
											VH.A(57265),
											VH.A(57369),
											VH.A(57376),
											VH.A(57387),
											VH.A(57400),
											VH.A(57409),
											VH.A(57416),
											VH.A(57429),
											VH.A(57448),
											VH.A(57469),
											VH.A(56560),
											VH.A(57484),
											VH.A(56571),
											VH.A(57503),
											VH.A(57522),
											VH.A(57535)
										};
										long num11 = array5.Length;
										object instance3 = NewLateBinding.LateGet(currentRegion.Cells[1, 4 + 2 * num6], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num11
										}, null, null, array4 = new bool[2] { false, true });
										if (array4[1])
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
											num11 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance3, null, VH.A(57302), new object[1] { array5 }, null, null, OptimisticSet: false, RValueBase: true);
										long num12 = num11 + 2 * num6 + 2;
										object instance4 = NewLateBinding.LateGet(currentRegion.Cells[1, num12 + 2], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num6
										}, null, null, array4 = new bool[2] { false, true });
										if (array4[1])
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
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance4, null, VH.A(57346), new object[1] { array }, null, null, OptimisticSet: false, RValueBase: true);
										object instance5 = NewLateBinding.LateGet(currentRegion.Cells[1, num12 + num6 + 2], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num6
										}, null, null, array4 = new bool[2] { false, true });
										if (array4[1])
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
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance5, null, VH.A(57346), new object[1] { array2 }, null, null, OptimisticSet: false, RValueBase: true);
										object instance6 = NewLateBinding.LateGet(currentRegion.Cells[1, num12 + 2 * num6 + 2], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num6
										}, null, null, array4 = new bool[2] { false, true });
										if (array4[1])
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
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance6, null, VH.A(57346), new object[1] { array }, null, null, OptimisticSet: false, RValueBase: true);
										object instance7 = NewLateBinding.LateGet(currentRegion.Cells[1, num12 + 3 * num6 + 2], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num6
										}, null, null, array4 = new bool[2] { false, true });
										if (array4[1])
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
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance7, null, VH.A(57346), new object[1] { array2 }, null, null, OptimisticSet: false, RValueBase: true);
										object instance8 = NewLateBinding.LateGet(currentRegion.Cells[1, num12 + 4 * num6 + 2], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num6
										}, null, null, array4 = new bool[2] { false, true });
										if (array4[1])
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
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance8, null, VH.A(57346), new object[1] { array }, null, null, OptimisticSet: false, RValueBase: true);
										object instance9 = NewLateBinding.LateGet(currentRegion.Cells[1, num12 + 5 * num6 + 2], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num6
										}, null, null, array4 = new bool[2] { false, true });
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
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance9, null, VH.A(57346), new object[1] { array }, null, null, OptimisticSet: false, RValueBase: true);
										object instance10 = NewLateBinding.LateGet(currentRegion.Cells[1, num12 + 6 * num6 + 2], null, VH.A(57333), array3 = new object[2]
										{
											Missing.Value,
											num6
										}, null, null, array4 = new bool[2] { false, true });
										if (array4[1])
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
											num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
										}
										NewLateBinding.LateSetComplex(instance10, null, VH.A(57346), new object[1] { array }, null, null, OptimisticSet: false, RValueBase: true);
										object instance11 = currentRegion.Cells[2, num6 + 2];
										NewLateBinding.LateSetComplex(instance11, null, VH.A(57346), new object[1] { VH.A(57546) + Conversions.ToString(currentRegion.Row + 2) + VH.A(57563) + Conversions.ToString(currentRegion.Row + 2) + VH.A(57572) + Conversions.ToString(1 + num6) + VH.A(57575) }, null, null, OptimisticSet: false, RValueBase: true);
										lD.A = Conversions.ToString(NewLateBinding.LateGet(instance11, null, VH.A(5814), new object[4]
										{
											true,
											true,
											XlReferenceStyle.xlR1C1,
											false
										}, null, null, null));
										mD.E.FormulaR1C1 = VH.A(57582) + lD.A + VH.A(57591);
										instance11 = null;
										object obj3 = NewLateBinding.LateGet(worksheet.Cells[currentRegion.Row, num12 + 7 * num6 + 3], null, VH.A(57333), array3 = new object[2] { num3, 2 }, null, null, array4 = new bool[2] { true, false });
										if (array4[0])
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
											num3 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(int));
										}
										Range range3 = (Range)obj3;
										long num13 = num9;
										for (long num14 = 1L; num14 <= num13; num14++)
										{
											NewLateBinding.LateSetComplex(range3.Cells[num14 + 1, 1], null, VH.A(57346), new object[1] { VH.A(57612) + Conversions.ToString(range2.Row + 3 * num14 - 1) + VH.A(57617) }, null, null, OptimisticSet: false, RValueBase: true);
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											NewLateBinding.LateSetComplex(range3.Cells[1, 2], null, VH.A(57302), new object[1] { VH.A(57622) }, null, null, OptimisticSet: false, RValueBase: true);
											object instance12 = range3.Cells[2, 1];
											string memberName2 = VH.A(57333);
											object[] obj4 = new object[1] { num9 };
											array3 = obj4;
											bool[] obj5 = new bool[1] { true };
											array4 = obj5;
											object obj6 = NewLateBinding.LateGet(instance12, null, memberName2, obj4, null, null, obj5);
											if (array4[0])
											{
												num9 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(long));
											}
											Range range4 = (Range)obj6;
											object obj7 = NewLateBinding.LateGet(range3.Cells[2, 2], null, VH.A(57333), array3 = new object[1] { num9 }, null, null, array4 = new bool[1] { true });
											if (array4[0])
											{
												num9 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[0]), typeof(long));
											}
											object instance13 = obj7;
											NewLateBinding.LateSetComplex(instance13, null, VH.A(57302), new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
											NewLateBinding.LateSetComplex(instance13, null, VH.A(57240), new object[1] { text3 }, null, null, OptimisticSet: false, RValueBase: true);
											instance13 = null;
											StackedWaterfall.A(range3);
											((_Worksheet)worksheet).get_Range((object)VH.A(57631), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize((object)15, (object)2).Font.Size = Operators.SubtractObject(((_Worksheet)worksheet).get_Range((object)VH.A(57631), RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Size, 1);
											string text4 = mD.A.get_Address((object)true, (object)true, XlReferenceStyle.xlR1C1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
											string text5 = mD.B.get_Address((object)true, (object)true, XlReferenceStyle.xlR1C1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
											string text6 = mD.C.get_Address((object)true, (object)true, XlReferenceStyle.xlR1C1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
											string text7 = mD.D.get_Address((object)true, (object)true, XlReferenceStyle.xlR1C1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
											string text8 = mD.E.get_Address((object)true, (object)true, XlReferenceStyle.xlR1C1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
											NewLateBinding.LateSetComplex(currentRegion.Cells[2, 1], null, VH.A(57302), new object[1] { VH.A(41385) }, null, null, OptimisticSet: false, RValueBase: true);
											NewLateBinding.LateSetComplex(currentRegion.Cells[2, 2 * num6 + 8], null, VH.A(57346), new object[1] { VH.A(57636) + lD.A + VH.A(57645) + Conversions.ToString(range2.Row + 2) + VH.A(57572) + Conversions.ToString(num6 + 2) + VH.A(57650) }, null, null, OptimisticSet: false, RValueBase: true);
											long num15 = num9;
											long num14 = 1L;
											while (num14 <= num15)
											{
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, 1], null, VH.A(57302), new object[1] { VH.A(41385) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 2, 1], null, VH.A(57302), new object[1] { VH.A(41385) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, num6 + 3], null, VH.A(57346), new object[1] { VH.A(57657) + Conversions.ToString(range2.Row) + VH.A(57688) + Conversions.ToString(num6 + 1) + VH.A(57723) }, null, null, OptimisticSet: false, RValueBase: true);
												object instance14 = NewLateBinding.LateGet(currentRegion.Cells[3 * num14, num6 + 4], null, VH.A(57333), array3 = new object[2]
												{
													Missing.Value,
													num6
												}, null, null, array4 = new bool[2] { false, true });
												if (array4[1])
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
													num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
												}
												NewLateBinding.LateSetComplex(instance14, null, VH.A(57346), new object[1] { VH.A(57732) + Conversions.ToString(num6 + 3) + VH.A(57745) + Conversions.ToString(range2.Row + 2) + VH.A(57758) + Conversions.ToString(num6 + 2) + VH.A(57765) + Conversions.ToString(num6 + 2) + VH.A(57778) + Conversions.ToString(num6 + 2) + VH.A(57807) + Conversions.ToString(num6 + 2) + VH.A(57822) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 4], null, VH.A(57346), new object[1] { VH.A(57732) + Conversions.ToString(num6 + 3) + VH.A(57833) + lD.A + VH.A(57745) + Conversions.ToString(range2.Row + 2) + VH.A(57758) + Conversions.ToString(num6 + 2) + VH.A(57765) + Conversions.ToString(num6 + 2) + VH.A(57842) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 5], null, VH.A(57346), new object[1] { VH.A(57859) + Conversions.ToString(num6 + 1) + VH.A(57878) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 6], null, VH.A(57346), new object[1] { VH.A(57897) + Conversions.ToString(num6 + 2) + VH.A(57926) + Conversions.ToString(num6 + 2) + VH.A(57981) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 2, 2 * num6 + 7], null, VH.A(57346), new object[1] { VH.A(58016) + Conversions.ToString(range2.Row) + VH.A(58041) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 8], null, VH.A(58046), new object[1] { VH.A(58071) + Conversions.ToString(num6 + 1) + VH.A(58114) + Conversions.ToString(num6 + 1) + VH.A(58133) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, 2 * num6 + 8], null, VH.A(58046), new object[1] { VH.A(58142) + Conversions.ToString(num6 + 1) + VH.A(58201) + Conversions.ToString(num6 + 1) + VH.A(58133) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 2, 2 * num6 + 8], null, VH.A(57346), new object[1] { VH.A(58236) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 9], null, VH.A(58046), new object[1] { VH.A(58251) + text8 + VH.A(58268) + Conversions.ToString(num6 + 3) + VH.A(58279) + Conversions.ToString(num6 + 3) + VH.A(58330) + Conversions.ToString(num6 + 4) + VH.A(58363) + Conversions.ToString(2 * num6 + 3) + VH.A(58370) + Conversions.ToString(num6 + 4) + VH.A(58363) + Conversions.ToString(2 * num6 + 3) + VH.A(58381) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, 2 * num6 + 9], null, VH.A(57346), new object[1] { VH.A(58236) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 2, 2 * num6 + 9], null, VH.A(57346), new object[1] { VH.A(58392) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 10], null, VH.A(57346), new object[1] { VH.A(58251) + text8 + VH.A(58268) + Conversions.ToString(num6 + 3) + VH.A(58407) + Conversions.ToString(num6 + 3) + VH.A(58450) + Conversions.ToString(num6 + 4) + VH.A(58363) + Conversions.ToString(2 * num6 + 3) + VH.A(58499) + Conversions.ToString(num6 + 4) + VH.A(58363) + Conversions.ToString(2 * num6 + 3) + VH.A(58530) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, 2 * num6 + 10], null, VH.A(57346), new object[1] { VH.A(58251) + text8 + VH.A(58537) + Conversions.ToString(num6 + 3) + VH.A(58556) + Conversions.ToString(num6 + 3) + VH.A(58593) + Conversions.ToString(num6 + 4) + VH.A(58650) + Conversions.ToString(2 * num6 + 3) + VH.A(58665) + Conversions.ToString(num6 + 4) + VH.A(58650) + Conversions.ToString(2 * num6 + 3) + VH.A(58700) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 2, 2 * num6 + 10], null, VH.A(57346), new object[1] { VH.A(58711) }, null, null, OptimisticSet: false, RValueBase: true);
												if (num14 < num9)
												{
													NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 2, 2 * num6 + 11], null, VH.A(57346), new object[1] { VH.A(57582) + text4 + VH.A(58726) + Conversions.ToString(num6 + 3) + VH.A(2378) + text8 + VH.A(58755) + Conversions.ToString(2 * num6 + 6) + VH.A(58790) }, null, null, OptimisticSet: false, RValueBase: true);
												}
												if (num14 < num9)
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
													NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 2, 2 * num6 + 12], null, VH.A(57346), new object[1] { VH.A(57582) + text4 + VH.A(58817) + Conversions.ToString(num6 + 3) + VH.A(58836) + Conversions.ToString(2 * num6 + 6) + VH.A(58790) }, null, null, OptimisticSet: false, RValueBase: true);
												}
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 13], null, VH.A(57302), new object[1] { (double)(3 * num14) - 0.5 }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 14], null, VH.A(57346), new object[1] { VH.A(58863) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 15], null, VH.A(57346), new object[1] { VH.A(57582) + text7 + VH.A(58878) + Conversions.ToString(2 * num6 + 5) + VH.A(58889) + text7 + VH.A(58906) + Conversions.ToString(2 * num6 + 5) + VH.A(58917) + Conversions.ToString(2 * num6 + 5) + VH.A(58932) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 16], null, VH.A(57346), new object[1] { VH.A(58943) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, 2 * num6 + 17], null, VH.A(57346), new object[1] { VH.A(57636) + text7 + VH.A(58988) + Conversions.ToString(2 * num6 + 5) + VH.A(59005) + Conversions.ToString(2 * num6 + 5) + VH.A(59016) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, num12], null, VH.A(57346), new object[1] { VH.A(59033) + Conversions.ToString(num6 + 3) + VH.A(2378) + text8 + VH.A(59054) + Conversions.ToString(2 * num6 + 5) + VH.A(57650) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, num12], null, VH.A(57346), new object[1] { VH.A(58236) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, num12 + 1], null, VH.A(57346), new object[1] { VH.A(59067) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, num12 + 1], null, VH.A(57346), new object[1] { VH.A(59067) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(currentRegion.Cells[3 * num14, num12 + 2], null, VH.A(57333), new object[2]
												{
													Missing.Value,
													num6 - 1
												}, null, null, null), null, VH.A(57346), new object[1] { VH.A(59033) + Conversions.ToString(num12) + VH.A(59156) + Conversions.ToString(16 + num6) + VH.A(59171) + Conversions.ToString(2 * num6 + 9) + VH.A(59192) + Conversions.ToString(num12 + 5) + VH.A(59219) + Conversions.ToString(16 + num6) + VH.A(59244) + Conversions.ToString(2 * num6 + 9) + VH.A(59192) + Conversions.ToString(num12 + 5) + VH.A(59253) }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, num12 + num6 + 1], null, VH.A(57346), new object[1] { VH.A(59033) + Conversions.ToString(num12) + VH.A(59156) + Conversions.ToString(16 + num6) + VH.A(59171) + Conversions.ToString(2 * num6 + 9) + VH.A(59270) + Conversions.ToString(16 + num6) + VH.A(59244) + Conversions.ToString(2 * num6 + 9) + VH.A(59293) }, null, null, OptimisticSet: false, RValueBase: true);
												object instance15 = NewLateBinding.LateGet(currentRegion.Cells[3 * num14 + 1, num12 + 2], null, VH.A(57333), array3 = new object[2]
												{
													Missing.Value,
													num6
												}, null, null, array4 = new bool[2] { false, true });
												if (array4[1])
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
													num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
												}
												NewLateBinding.LateSetComplex(instance15, null, VH.A(57346), new object[1] { VH.A(59033) + Conversions.ToString(num12) + VH.A(59308) + Conversions.ToString(16 + num6) + VH.A(59331) + Conversions.ToString(2 * num6 + 6) + VH.A(59360) }, null, null, OptimisticSet: false, RValueBase: true);
												long num16 = num6;
												for (long num17 = 1L; num17 <= num16; num17++)
												{
													NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, num12 + num6 + num17 + 1], null, VH.A(57346), new object[1] { VH.A(59033) + Conversions.ToString(num12) + VH.A(59156) + Conversions.ToString(2 * num17 + num6 + 15) + VH.A(59387) + Conversions.ToString(2 * num17 - 1) + VH.A(59406) + Conversions.ToString(2 * num17 + num6 + 15) + VH.A(59419) }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, num12 + num6 + num17 + 1], null, VH.A(57346), new object[1] { VH.A(59428) + Conversions.ToString(num12) + VH.A(59308) + Conversions.ToString(2 * num17 + num6 + 15) + VH.A(59331) + Conversions.ToString(2 * num6 + 6) + VH.A(59360) }, null, null, OptimisticSet: false, RValueBase: true);
												}
												while (true)
												{
													switch (5)
													{
													case 0:
														continue;
													}
													object instance16 = currentRegion.Cells[3 * num14, num12 + 2 * num6 + 2];
													string memberName3 = VH.A(57333);
													object[] obj8 = new object[2]
													{
														Missing.Value,
														num6
													};
													array3 = obj8;
													bool[] obj9 = new bool[2] { false, true };
													array4 = obj9;
													object instance17 = NewLateBinding.LateGet(instance16, null, memberName3, obj8, null, null, obj9);
													if (array4[1])
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
														num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
													}
													NewLateBinding.LateSetComplex(instance17, null, VH.A(57346), new object[1] { VH.A(59033) + Conversions.ToString(num12) + VH.A(59156) + Conversions.ToString(3 * num6 + 16) + VH.A(59457) + Conversions.ToString(2 * num6 + 6) + VH.A(59478) }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(currentRegion.Cells[3 * num14 + 1, num12 + 2 + num6 * 2], null, VH.A(57333), new object[2]
													{
														Missing.Value,
														num6 - 1
													}, null, null, null), null, VH.A(57346), new object[1] { VH.A(59428) + Conversions.ToString(num12) + VH.A(59308) + Conversions.ToString(3 * num6 + 16) + VH.A(59457) + Conversions.ToString(2 * num6 + 9) + VH.A(59192) + Conversions.ToString(num12 + 3 * num6 + 1) + VH.A(59503) + Conversions.ToString(3 * num6 + 16) + VH.A(59244) + Conversions.ToString(2 * num6 + 9) + VH.A(59192) + Conversions.ToString(num12 + 3 * num6 + 1) + VH.A(59253) }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, num12 + 3 * num6 + 1], null, VH.A(57346), new object[1] { VH.A(59428) + Conversions.ToString(num12) + VH.A(59308) + Conversions.ToString(3 * num6 + 16) + VH.A(59457) + Conversions.ToString(2 * num6 + 9) + VH.A(59538) + Conversions.ToString(3 * num6 + 16) + VH.A(59244) + Conversions.ToString(2 * num6 + 9) + VH.A(59293) }, null, null, OptimisticSet: false, RValueBase: true);
													long num18 = num6;
													for (long num19 = 1L; num19 <= num18; num19++)
													{
														NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14, num12 + 3 * num6 + num19 + 1], null, VH.A(57346), new object[1] { VH.A(59033) + Conversions.ToString(num12) + VH.A(59156) + Conversions.ToString(3 * num6 + 2 * num19 + 15) + VH.A(59457) + Conversions.ToString(2 * num6 + 6) + VH.A(59478) }, null, null, OptimisticSet: false, RValueBase: true);
														NewLateBinding.LateSetComplex(currentRegion.Cells[3 * num14 + 1, num12 + 3 * num6 + num19 + 1], null, VH.A(57346), new object[1] { VH.A(59428) + Conversions.ToString(num12) + VH.A(59308) + Conversions.ToString(3 * num6 + 2 * num19 + 15) + VH.A(59571) + Conversions.ToString(3 * num6 + 2 * num19 + 15) + VH.A(59598) + Conversions.ToString(2 * num19 - 1) + VH.A(59419) }, null, null, OptimisticSet: false, RValueBase: true);
													}
													while (true)
													{
														switch (7)
														{
														case 0:
															continue;
														}
														object instance18 = currentRegion.Cells[3 * num14, num12 + 4 * num6 + 2];
														string memberName4 = VH.A(57333);
														object[] obj10 = new object[2]
														{
															Missing.Value,
															num6
														};
														array3 = obj10;
														bool[] obj11 = new bool[2] { false, true };
														array4 = obj11;
														object instance19 = NewLateBinding.LateGet(instance18, null, memberName4, obj10, null, null, obj11);
														if (array4[1])
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
															num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
														}
														NewLateBinding.LateSetComplex(instance19, null, VH.A(57346), new object[1] { VH.A(59611) + Conversions.ToString(2 * num6 + 13) + VH.A(59618) + Conversions.ToString(2 * num6 + 6) + VH.A(59625) }, null, null, OptimisticSet: false, RValueBase: true);
														object instance20 = NewLateBinding.LateGet(currentRegion.Cells[3 * num14 + 1, num12 + 4 * num6 + 2], null, VH.A(57333), array3 = new object[2]
														{
															Missing.Value,
															num6
														}, null, null, array4 = new bool[2] { false, true });
														if (array4[1])
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
															num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
														}
														NewLateBinding.LateSetComplex(instance20, null, VH.A(57346), new object[1] { VH.A(58236) + Conversions.ToString(2 * num6 + 13) + VH.A(59630) + Conversions.ToString(2 * num6 + 6) + VH.A(59625) }, null, null, OptimisticSet: false, RValueBase: true);
														object instance21 = NewLateBinding.LateGet(currentRegion.Cells[3 * num14, num12 + 5 * num6 + 2], null, VH.A(57333), array3 = new object[2]
														{
															Missing.Value,
															num6
														}, null, null, array4 = new bool[2] { false, true });
														if (array4[1])
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
															num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
														}
														NewLateBinding.LateSetComplex(instance21, null, VH.A(57346), new object[1] { VH.A(59645) + Conversions.ToString(6 * num6 + 16) + VH.A(59670) + Conversions.ToString(num12) + VH.A(59685) + Conversions.ToString(num6 + 3) + VH.A(59704) + Conversions.ToString(2 * num6 + 8) + VH.A(59723) + Conversions.ToString(num6 + 4) + VH.A(59746) + Conversions.ToString(6 * num6 + 16) + VH.A(59244) + Conversions.ToString(num6 + 4) + VH.A(59746) + Conversions.ToString(6 * num6 + 16) + VH.A(59757) + Conversions.ToString(6 * num6 + 16) + VH.A(59798) }, null, null, OptimisticSet: false, RValueBase: true);
														object instance22 = NewLateBinding.LateGet(currentRegion.Cells[3 * num14 + 1, num12 + 5 * num6 + 2], null, VH.A(57333), array3 = new object[2]
														{
															Missing.Value,
															num6
														}, null, null, array4 = new bool[2] { false, true });
														if (array4[1])
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
															num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
														}
														NewLateBinding.LateSetComplex(instance22, null, VH.A(57346), new object[1] { VH.A(59821) + Conversions.ToString(6 * num6 + 16) + VH.A(59852) + text8 + VH.A(58537) + Conversions.ToString(num6 + 3) + VH.A(59869) + Conversions.ToString(2 * num6 + 8) + VH.A(59890) + Conversions.ToString(num6 + 4) + VH.A(59919) + Conversions.ToString(6 * num6 + 16) + VH.A(59938) + Conversions.ToString(num6 + 4) + VH.A(59919) + Conversions.ToString(6 * num6 + 16) + VH.A(59955) + Conversions.ToString(6 * num6 + 16) + VH.A(60004) }, null, null, OptimisticSet: false, RValueBase: true);
														object instance23 = NewLateBinding.LateGet(currentRegion.Cells[3 * num14, num12 + 6 * num6 + 2], null, VH.A(57333), array3 = new object[2]
														{
															Missing.Value,
															num6
														}, null, null, array4 = new bool[2] { false, true });
														if (array4[1])
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
															num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
														}
														NewLateBinding.LateSetComplex(instance23, null, VH.A(57346), new object[1] { VH.A(58251) + text5 + VH.A(60023) + Conversions.ToString(7 * num6 + 16) + VH.A(60034) + text6 + VH.A(60041) + Conversions.ToString(num12) + VH.A(60048) + Conversions.ToString(7 * num6 + 16) + VH.A(60065) }, null, null, OptimisticSet: false, RValueBase: true);
														object instance24 = NewLateBinding.LateGet(currentRegion.Cells[3 * num14 + 1, num12 + 6 * num6 + 2], null, VH.A(57333), array3 = new object[2]
														{
															Missing.Value,
															num6
														}, null, null, array4 = new bool[2] { false, true });
														if (array4[1])
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
															num6 = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array3[1]), typeof(long));
														}
														NewLateBinding.LateSetComplex(instance24, null, VH.A(57346), new object[1] { VH.A(60076) + text5 + VH.A(60093) + Conversions.ToString(7 * num6 + 16) + VH.A(60112) + text6 + VH.A(60121) + Conversions.ToString(num12) + VH.A(60128) + Conversions.ToString(7 * num6 + 16) + VH.A(60065) }, null, null, OptimisticSet: false, RValueBase: true);
														num14++;
														break;
													}
													break;
												}
											}
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												Range range5 = currentRegion.get_Offset((object)1, RuntimeHelpers.GetObjectValue(Missing.Value)).get_Resize((object)(3 * num9 + 2), RuntimeHelpers.GetObjectValue(Missing.Value));
												range5.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)1).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(num6 + 1)).NumberFormat = text3;
												range5.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(num6 + 3)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(num6 + 2)).NumberFormat = text3;
												range5.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(2 * num6 + 7)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)3).NumberFormat = text3;
												range5.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(2 * num6 + 13)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)4).NumberFormat = text3;
												range5.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(2 * num6 + 17)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(4 * num6 + 2)).NumberFormat = text3;
												range5.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(7 * num6 + 19)).get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)(2 * num6)).NumberFormat = text3;
												_ = null;
												NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(60153), new object[0], null, null, null, IgnoreReturn: true);
												Range arg = (Range)NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 9, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(57333), new object[2]
												{
													Missing.Value,
													2
												}, null, null, null);
												ND nD = StackedWaterfall.A(application.WorksheetFunction.Max(application.WorksheetFunction.Max(arg, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), application.WorksheetFunction.Min(application.WorksheetFunction.Min(arg, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
												mD.C.Value2 = nD.C / 2.0;
												Range source = application.Union((Range)NewLateBinding.LateGet(currentRegion.Cells[1, 1], null, VH.A(57333), new object[1] { 3 * num9 + 2 }, null, null, null), (Range)NewLateBinding.LateGet(currentRegion.Cells[1, 2 * num6 + 18], null, VH.A(57333), new object[2]
												{
													3 * num9 + 2,
													4 * num6 + 2
												}, null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
												Chart chart2 = chart;
												chart2.SetSourceData(source, XlSearchOrder.xlByColumns);
												NewLateBinding.LateSetComplex(chart2.ChartGroups(1), null, VH.A(60172), new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
												object instance25 = chart2.SeriesCollection(1);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance25, null, VH.A(60189), new object[0], null, null, null), null, VH.A(60202), new object[1] { Microsoft.Office.Interop.Excel.Constants.xlNone }, null, null, OptimisticSet: false, RValueBase: true);
												object instance26 = NewLateBinding.LateGet(NewLateBinding.LateGet(instance25, null, VH.A(60221), new object[0], null, null, null), null, VH.A(60234), new object[0], null, null, null);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance26, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60262), new object[1] { MsoThemeColorIndex.msoThemeColorLight1 }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance26, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60295), new object[1] { -0.35 }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance26, null, VH.A(60316), new object[0], null, null, null), null, VH.A(60262), new object[1] { MsoThemeColorIndex.msoThemeColorLight1 }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance26, null, VH.A(60316), new object[0], null, null, null), null, VH.A(60295), new object[1] { -0.15 }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateCall(instance26, null, VH.A(60335), new object[1] { MsoPatternType.msoPattern70Percent }, null, null, null, IgnoreReturn: true);
												instance26 = null;
												instance25 = null;
												object instance27 = chart2.SeriesCollection(2);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance27, null, VH.A(60189), new object[0], null, null, null), null, VH.A(60202), new object[1] { Microsoft.Office.Interop.Excel.Constants.xlNone }, null, null, OptimisticSet: false, RValueBase: true);
												NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance27, null, VH.A(36170), new object[0], null, null, null), null, VH.A(60354), new object[1] { Microsoft.Office.Interop.Excel.Constants.xlNone }, null, null, OptimisticSet: false, RValueBase: true);
												instance27 = null;
												long num20 = num6;
												for (long num21 = 1L; num21 <= num20; num21++)
												{
													MsoThemeColorIndex msoThemeColorIndex = StackedWaterfall.A(num21);
													double num22 = StackedWaterfall.A(num21);
													object instance28 = NewLateBinding.LateGet(NewLateBinding.LateGet(chart2.SeriesCollection(2 + num21), null, VH.A(60221), new object[0], null, null, null), null, VH.A(60234), new object[0], null, null, null);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance28, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60262), new object[1] { msoThemeColorIndex }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance28, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60295), new object[1] { num22 }, null, null, OptimisticSet: false, RValueBase: true);
													instance28 = null;
													object instance29 = NewLateBinding.LateGet(NewLateBinding.LateGet(chart2.SeriesCollection(3 + 2 * num6 - num21), null, VH.A(60221), new object[0], null, null, null), null, VH.A(60234), new object[0], null, null, null);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance29, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60262), new object[1] { msoThemeColorIndex }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance29, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60295), new object[1] { num22 }, null, null, OptimisticSet: false, RValueBase: true);
													instance29 = null;
													object instance30 = NewLateBinding.LateGet(NewLateBinding.LateGet(chart2.SeriesCollection(2 + 2 * num6 + num21), null, VH.A(60221), new object[0], null, null, null), null, VH.A(60234), new object[0], null, null, null);
													NewLateBinding.LateCall(instance30, null, VH.A(60335), new object[1] { MsoPatternType.msoPatternDarkDownwardDiagonal }, null, null, null, IgnoreReturn: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance30, null, VH.A(60316), new object[0], null, null, null), null, VH.A(60262), new object[1] { msoThemeColorIndex }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance30, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60262), new object[1] { msoThemeColorIndex }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance30, null, VH.A(60316), new object[0], null, null, null), null, VH.A(60295), new object[1] { num22 + 0.4 }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance30, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60295), new object[1] { num22 }, null, null, OptimisticSet: false, RValueBase: true);
													instance30 = null;
													object instance31 = NewLateBinding.LateGet(NewLateBinding.LateGet(chart2.SeriesCollection(3 + 4 * num6 - num21), null, VH.A(60221), new object[0], null, null, null), null, VH.A(60234), new object[0], null, null, null);
													NewLateBinding.LateCall(instance31, null, VH.A(60335), new object[1] { MsoPatternType.msoPatternDarkDownwardDiagonal }, null, null, null, IgnoreReturn: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance31, null, VH.A(60316), new object[0], null, null, null), null, VH.A(60262), new object[1] { msoThemeColorIndex }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance31, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60262), new object[1] { msoThemeColorIndex }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance31, null, VH.A(60316), new object[0], null, null, null), null, VH.A(60295), new object[1] { num22 + 0.4 }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance31, null, VH.A(60243), new object[0], null, null, null), null, VH.A(60295), new object[1] { num22 }, null, null, OptimisticSet: false, RValueBase: true);
													instance31 = null;
												}
												while (true)
												{
													switch (5)
													{
													case 0:
														continue;
													}
													PlotArea plotArea = chart2.PlotArea;
													plotArea.Format.Line.Visible = MsoTriState.msoFalse;
													plotArea.Format.Fill.Visible = MsoTriState.msoFalse;
													_ = null;
													object instance32 = chart2.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
													NewLateBinding.LateSetComplex(instance32, null, VH.A(60375), new object[1] { XlCategoryType.xlCategoryScale }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance32, null, VH.A(60400), new object[0], null, null, null), null, VH.A(60421), new object[1] { XlTickLabelOrientation.xlTickLabelOrientationHorizontal }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(instance32, null, VH.A(60444), new object[1] { 1 }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(instance32, null, VH.A(60477), new object[1] { Microsoft.Office.Interop.Excel.Constants.xlNone }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(instance32, null, VH.A(60512), new object[1] { XlTickMark.xlTickMarkNone }, null, null, OptimisticSet: false, RValueBase: true);
													NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance32, null, VH.A(60221), new object[0], null, null, null), null, VH.A(57400), new object[0], null, null, null), null, VH.A(60243), new object[0], null, null, null), null, VH.A(60539), new object[1] { 10921638L }, null, null, OptimisticSet: false, RValueBase: true);
													instance32 = null;
													Series series = (Series)NewLateBinding.LateGet(chart2.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60546), new object[0], null, null, null);
													Series series2 = series;
													series2.Name = Conversions.ToString(Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(currentRegion.Cells[1, 2 * num6 + 8], null, VH.A(5814), new object[4]
													{
														Missing.Value,
														Missing.Value,
														Missing.Value,
														true
													}, null, null, null)));
													series2.Values = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 8, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
													series2.XValues = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 7, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
													series2.ChartType = XlChartType.xlXYScatter;
													series2.AxisGroup = XlAxisGroup.xlPrimary;
													series2.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
													series2.ErrorBar(XlErrorBarDirection.xlX, XlErrorBarInclude.xlErrorBarIncludeBoth, XlErrorBarType.xlErrorBarTypeCustom, VH.A(48936) + StackedWaterfall.A((Range)NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 12, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null)), VH.A(48936) + StackedWaterfall.A((Range)NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 11, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null)));
													series2.ErrorBars.EndStyle = XlEndStyleCap.xlNoCap;
													series2.ErrorBars.Format.Line.ForeColor.RGB = lD.A;
													series2 = null;
													long num23 = 1L;
													do
													{
														series = (Series)NewLateBinding.LateGet(chart2.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60546), new object[0], null, null, null);
														Series series3 = series;
														series3.Name = Conversions.ToString(Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(currentRegion.Cells[1, 2 * num6 + 12 + 2 * num23], null, VH.A(5814), new object[4]
														{
															Missing.Value,
															Missing.Value,
															Missing.Value,
															true
														}, null, null, null)));
														series3.Values = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 12 + 2 * num23, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
														series3.XValues = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 13, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
														series3.ChartType = XlChartType.xlXYScatter;
														series3.AxisGroup = XlAxisGroup.xlPrimary;
														series3.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
														series3.Border.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
														series3.HasDataLabels = true;
														NewLateBinding.LateSetComplex(series3.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60578), new object[1] { Interaction.IIf(num23 == 1, XlDataLabelPosition.xlLabelPositionAbove, XlDataLabelPosition.xlLabelPositionBelow) }, null, null, OptimisticSet: false, RValueBase: true);
														series3 = null;
														StackedWaterfall.A(series, (Range)NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[2 * num6 + 13 + 2 * num23, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
														num23++;
													}
													while (num23 <= 2);
													while (true)
													{
														switch (6)
														{
														case 0:
															continue;
														}
														long num24 = num6;
														for (long num25 = 1L; num25 <= num24; num25++)
														{
															series = (Series)NewLateBinding.LateGet(chart2.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60546), new object[0], null, null, null);
															Series series4 = series;
															series4.Name = Conversions.ToString(Operators.ConcatenateObject(VH.A(48936), NewLateBinding.LateGet(currentRegion.Cells[1, num12 + 5 * num6 + 1 + num25], null, VH.A(5814), new object[4]
															{
																Missing.Value,
																Missing.Value,
																Missing.Value,
																true
															}, null, null, null)));
															series4.Values = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[num12 + 5 * num6 + 1 + num25, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
															series4.XValues = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[num12 + 4 * num6 + 1 + num25, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
															series4.ChartType = XlChartType.xlXYScatter;
															series4.AxisGroup = XlAxisGroup.xlPrimary;
															series4.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
															series4.Border.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
															series4.HasDataLabels = true;
															NewLateBinding.LateSetComplex(series4.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60578), new object[1] { XlDataLabelPosition.xlLabelPositionCenter }, null, null, OptimisticSet: false, RValueBase: true);
															series4 = null;
															StackedWaterfall.A(series, (Range)NewLateBinding.LateGet(NewLateBinding.LateGet(currentRegion.Columns[num12 + 6 * num6 + 1 + num25, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { 1 }, null, null, null), null, VH.A(57333), new object[1] { 3 * num9 + 1 }, null, null, null));
															object instance33 = NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60221), new object[0], null, null, null), null, VH.A(60595), new object[0], null, null, null), null, VH.A(60616), new object[0], null, null, null), null, VH.A(60635), new object[0], null, null, null);
															NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance33, null, VH.A(60234), new object[0], null, null, null), null, VH.A(60243), new object[0], null, null, null), null, VH.A(60262), new object[1] { MsoThemeColorIndex.msoThemeColorBackground1 }, null, null, OptimisticSet: false, RValueBase: true);
															object instance34 = NewLateBinding.LateGet(instance33, null, VH.A(60644), new object[0], null, null, null);
															NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance34, null, VH.A(55331), new object[0], null, null, null), null, VH.A(60262), new object[1] { MsoThemeColorIndex.msoThemeColorBackground1 }, null, null, OptimisticSet: false, RValueBase: true);
															NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance34, null, VH.A(55331), new object[0], null, null, null), null, VH.A(60295), new object[1] { -0.5 }, null, null, OptimisticSet: false, RValueBase: true);
															NewLateBinding.LateSetComplex(instance34, null, VH.A(60653), new object[1] { 0.6 }, null, null, OptimisticSet: false, RValueBase: true);
															NewLateBinding.LateSetComplex(instance34, null, VH.A(60678), new object[1] { 5 }, null, null, OptimisticSet: false, RValueBase: true);
															instance34 = null;
															instance33 = null;
														}
														while (true)
														{
															switch (6)
															{
															case 0:
																continue;
															}
															series = (Series)NewLateBinding.LateGet(chart2.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60546), new object[0], null, null, null);
															Series series5 = series;
															series5.Name = VH.A(48936) + range4.get_Offset((object)(-1), (object)1).get_Resize((object)1, (object)1).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
															series5.Values = range4.get_Offset((object)0, (object)1);
															series5.XValues = range4;
															series5.ChartType = XlChartType.xlLine;
															series5.AxisGroup = XlAxisGroup.xlSecondary;
															series5.MarkerStyle = XlMarkerStyle.xlMarkerStyleNone;
															series5.Border.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
															_ = null;
															chart2.HasLegend = false;
															chart2.HasLegend = true;
															chart2.Legend.Position = XlLegendPosition.xlLegendPositionRight;
															chart2.Legend.Border.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
															long num26 = 6 + num6;
															for (long num27 = 1L; num27 <= num26; num27++)
															{
																NewLateBinding.LateCall(chart2.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(chart2.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(52690), new object[0], null, null, null))), null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
															}
															while (true)
															{
																switch (1)
																{
																case 0:
																	continue;
																}
																long num28 = 3 * num6;
																for (long num29 = 1L; num29 <= num28; num29++)
																{
																	NewLateBinding.LateCall(chart2.Legend.LegendEntries(1), null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
																}
																object instance35 = chart2.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory);
																NewLateBinding.LateSetComplex(instance35, null, VH.A(60477), new object[1] { Microsoft.Office.Interop.Excel.Constants.xlNone }, null, null, OptimisticSet: false, RValueBase: true);
																NewLateBinding.LateSetComplex(instance35, null, VH.A(60512), new object[1] { Microsoft.Office.Interop.Excel.Constants.xlNone }, null, null, OptimisticSet: false, RValueBase: true);
																NewLateBinding.LateSetComplex(instance35, null, VH.A(60704), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
																instance35 = null;
																((_Chart)chart2).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlValue, (object)XlAxisGroup.xlSecondary, (object)false);
																((_Chart)chart2).set_HasAxis((object)Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, (object)XlAxisGroup.xlSecondary, (object)true);
																object instance36 = chart2.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, XlAxisGroup.xlSecondary);
																NewLateBinding.LateSetComplex(instance36, null, VH.A(60477), new object[1] { XlTickLabelPosition.xlTickLabelPositionLow }, null, null, OptimisticSet: false, RValueBase: true);
																NewLateBinding.LateSetComplex(NewLateBinding.LateGet(instance36, null, VH.A(60400), new object[0], null, null, null), null, VH.A(60565), new object[1] { 0 }, null, null, OptimisticSet: false, RValueBase: true);
																NewLateBinding.LateSetComplex(NewLateBinding.LateGet(NewLateBinding.LateGet(instance36, null, VH.A(60221), new object[0], null, null, null), null, VH.A(57400), new object[0], null, null, null), null, VH.A(41367), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
																instance36 = null;
																NewLateBinding.LateCall(currentRegion.Cells[1, 1], null, VH.A(51162), new object[0], null, null, null, IgnoreReturn: true);
																chart2 = null;
																break;
															}
															break;
														}
														break;
													}
													break;
												}
												break;
											}
											break;
										}
										break;
									}
									break;
								}
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								Forms.ErrorMessage(ex2.Message);
								clsReporting.LogException(ex2);
								ProjectData.ClearProjectError();
							}
							QuickCharts2.RestoreExcel(application, calc);
							calc = (XlCalculation)0;
							worksheet = null;
							chartObject = null;
							chart = null;
							lD = default(LD);
							mD = default(MD);
							QuickCharts2.LogActivity(VH.A(60747));
						}
					}
					A = null;
				}
			}
			application = null;
		}
	}

	private static ListBox A(Range A, string B, string C, string D)
	{
		Range range = A;
		ListBox obj = (ListBox)NewLateBinding.LateGet(NewLateBinding.LateGet(MH.A.Application.ActiveSheet, null, VH.A(60794), new object[0], null, null, null), null, VH.A(60813), new object[4] { range.Left, range.Top, range.Width, range.Height }, null, null, null);
		range = null;
		obj.Height = Conversions.ToDouble(A.Height);
		obj.Name = D;
		obj.ListFillRange = B;
		obj.MultiSelect = -4142;
		obj.ListIndex = 1;
		obj.LinkedCell = C;
		_ = null;
		return obj;
	}

	private static ND A(double A, double B, double C, double D, double E, bool F = false, bool G = false)
	{
		double num = D * Math.Pow(10.0, Conversion.Int(C));
		double d = E * Math.Pow(10.0, Conversion.Int(C));
		ND result = default(ND);
		if (F)
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
			result.A = A;
		}
		else
		{
			result.A = num * Conversion.Int(A / num);
		}
		if (G)
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
			result.B = B;
		}
		else
		{
			result.B = num * (Conversion.Int(B / num) + 1.0);
		}
		result.C = num;
		result.D = d;
		return result;
	}

	private static ND A(double A, double B, bool C = false, bool D = false, double E = 0.0, double F = 0.0)
	{
		if (B < A)
		{
			double num = B;
			B = A;
			A = num;
		}
		if (B == A)
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
			if (B > 0.0)
			{
				if (!D)
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
					B *= 1.01;
				}
				if (!C)
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
					A = 0.0;
				}
			}
			else if (B < 0.0)
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
				if (!C)
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
					A *= 1.01;
				}
				if (!D)
				{
					B = 0.0;
				}
			}
			else if (!D)
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
				B = 1.0;
			}
			if (B == A)
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
				if (!(B > 0.0))
				{
					B = ((!(B < 0.0)) ? 1.0 : 0.0);
				}
				else
				{
					A = 0.0;
				}
			}
		}
		else
		{
			if (!D)
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
				if (B > 0.0)
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
					B += (B - A) * 0.01;
				}
				else if (B < 0.0)
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
					B = MH.A.Application.WorksheetFunction.Min(B + (B - A) * 0.5, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				else
				{
					B = 0.0;
				}
			}
			if (!C)
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
				if (A > 0.0)
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
					A = MH.A.Application.WorksheetFunction.Max(A - (B - A) * 0.5, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				else if (A < 0.0)
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
					A -= (B - A) * 0.01;
				}
				else
				{
					A = 0.0;
				}
			}
		}
		if (A > 0.0)
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
			if (!C)
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
				if (A / B < 0.25)
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
					A = 0.0;
				}
			}
		}
		else if (B < 0.0 && !D)
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
			if (B / A < 0.25)
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
				B = 0.0;
			}
		}
		double num2 = Math.Log(B - A) / Math.Log(10.0);
		double num3 = Math.Pow(10.0, num2 - Conversion.Int(num2));
		double num4;
		double e;
		if (num3 >= 0.0)
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
			if (num3 <= 2.5)
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
				num4 = 0.2;
				e = 0.05;
				goto IL_0670;
			}
		}
		if (num3 >= 2.5)
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
			if (num3 <= 5.0)
			{
				num4 = 0.5;
				e = 0.1;
				goto IL_0670;
			}
		}
		if (num3 >= 5.0 && num3 <= 7.5)
		{
			num4 = 1.0;
			e = 0.2;
		}
		else
		{
			num4 = 2.0;
			e = 0.5;
		}
		goto IL_0670;
		IL_0670:
		ND result = StackedWaterfall.A(A, B, num2, num4, e, C, D);
		if (E > 0.0 && F > 0.0)
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
			long num5 = checked((long)Math.Round((result.B - result.A) / result.C));
			double num6 = E / (double)num5 / F;
			if (num6 < 1.35)
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
				double num7 = num4;
				if (num7 == 0.2)
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
					num4 = 0.5;
					e = 0.1;
				}
				else if (num7 == 0.5)
				{
					num4 = 1.0;
					e = 0.2;
				}
				else if (num7 == 1.0)
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
					num4 = 2.0;
					e = 0.5;
				}
				else if (num7 == 2.0)
				{
					num4 = 5.0;
					e = 1.0;
				}
				result = StackedWaterfall.A(A, B, num2, num4, e, C, D);
			}
			else if (num6 > 6.0)
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
				double num8 = num4;
				if (num8 == 0.2)
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
					num4 = 0.1;
					e = 0.02;
				}
				else if (num8 == 0.5)
				{
					num4 = 0.2;
					e = 0.05;
				}
				else if (num8 == 1.0)
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
					num4 = 0.5;
					e = 0.1;
				}
				else if (num8 == 2.0)
				{
					num4 = 1.0;
					e = 0.2;
				}
				result = StackedWaterfall.A(A, B, num2, num4, e, C, D);
			}
		}
		return result;
	}

	private static CheckBox A(Range A, string B, string C, string D, bool E = false, bool F = false)
	{
		double num = 0.0;
		if (F)
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
			num = 0.75;
		}
		double num2 = 0.0;
		double num3 = 0.0;
		Range range = A;
		num2 = Conversions.ToDouble(range.Height);
		num3 = Conversions.ToDouble(range.Width);
		CheckBox obj = (CheckBox)NewLateBinding.LateGet(NewLateBinding.LateGet(MH.A.Application.ActiveSheet, null, VH.A(60820), new object[0], null, null, null), null, VH.A(60813), new object[4]
		{
			Operators.AddObject(range.Left, num),
			Operators.AddObject(range.Top, num),
			num3 - 2.0 * num,
			num2 - 2.0 * num
		}, null, null, null);
		range = null;
		CheckBox checkBox = obj;
		checkBox.Name = C;
		checkBox.Caption = D;
		checkBox.LinkedCell = B;
		checkBox.Value = true;
		if (E)
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
			checkBox.Interior.Color = Microsoft.Office.Interop.Excel.Constants.xlNone;
		}
		else
		{
			checkBox.Interior.Color = XlRgbColor.rgbWhite;
		}
		checkBox.Width = num3 - 2.0 * num;
		checkBox.Height = num2 - 2.0 * num;
		checkBox = null;
		return obj;
	}

	private static bool A(string A, [Optional][DefaultParameterValue(false)] ref bool B, [Optional][DefaultParameterValue(false)] ref bool C)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		bool flag = true;
		Range range = null;
		try
		{
			range = ((_Application)application).get_Range((object)A, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (range != null)
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
			int count = range.Rows.Count;
			int count2 = range.Columns.Count;
			if (range.Areas.Count > 1)
			{
				flag = false;
			}
			else if (count < 3)
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
				flag = false;
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
				if (Strings.Len(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range[1, 1], null, VH.A(41636), new object[0], null, null, null))) == 0)
				{
					B = true;
				}
				else
				{
					long num = count2;
					long num2 = 2L;
					while (true)
					{
						if (num2 <= num)
						{
							if (((0 - ((!Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range[1, num2], null, VH.A(41636), new object[0], null, null, null)))) ? 1 : 0)) & Strings.Len(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range[1, num2], null, VH.A(41636), new object[0], null, null, null)))) != 0)
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
								B = true;
								break;
							}
							num2 = checked(num2 + 1);
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
				int num3 = Conversions.ToInteger(Interaction.IIf(B, 2, 1));
				bool flag2 = false;
				if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range[num3, count2], null, VH.A(41636), new object[0], null, null, null))))
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
					flag2 = true;
				}
				bool flag3;
				checked
				{
					int num4 = count2 - 1;
					int num5 = 2;
					while (true)
					{
						if (num5 <= num4)
						{
							if (Strings.Len(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range[num3, num5], null, VH.A(41636), new object[0], null, null, null))) != 0)
							{
								break;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0264;
								}
								continue;
								end_IL_0264:
								break;
							}
							num5++;
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
					flag3 = true;
					long num6 = num3 + 1;
					long num7 = count;
					long num8 = num6;
					while (true)
					{
						if (num8 <= num7)
						{
							if (Strings.Len(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range[num8, count2], null, VH.A(41636), new object[0], null, null, null))) != 0)
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
								flag3 = false;
								break;
							}
							num8++;
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
				}
				C = flag2 && flag3;
			}
		}
		else
		{
			flag = false;
		}
		return flag;
	}

	private static LD A(ref Range A, ref bool B)
	{
		XmlDocument xmlDoc = KH.A.SettingsXml;
		LD result = default(LD);
		wpfStackedWaterfall wpfStackedWaterfall2 = new wpfStackedWaterfall();
		QuickCharts2.HandleColorChange(wpfStackedWaterfall2.btnBridgeLine);
		wpfStackedWaterfall2.Range = A;
		wpfStackedWaterfall2.txtAddress.Text = A.get_Address((object)1, (object)1, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		QuickCharts2.LoadCommonSettings(xmlDoc, wpfStackedWaterfall2.numChartWidth, wpfStackedWaterfall2.numChartHeight);
		wpfStackedWaterfall2.chkPreserveFormulas.IsChecked = QuickCharts2.GetPreserveFormulas(xmlDoc);
		JD jD = (JD)Conversions.ToInteger(xmlDoc.SelectSingleNode(StackedWaterfall.m_A + VH.A(55388)).InnerText);
		if (jD != JD.A)
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
			if (jD != JD.B)
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
				wpfStackedWaterfall2.radLabelNone.IsChecked = true;
			}
			else
			{
				wpfStackedWaterfall2.radLabelAbove.IsChecked = true;
			}
		}
		else
		{
			wpfStackedWaterfall2.radLabelAboveBelow.IsChecked = true;
		}
		QuickCharts2.SetButtonColor(wpfStackedWaterfall2.btnBridgeLine, xmlDoc.SelectSingleNode(StackedWaterfall.m_A + VH.A(60841)).InnerText);
		wpfStackedWaterfall2.chkErrorBars.IsChecked = Conversions.ToBoolean(xmlDoc.SelectSingleNode(StackedWaterfall.m_A + VH.A(60864)).InnerText);
		wpfStackedWaterfall2.ShowDialog();
		if (wpfStackedWaterfall2.DialogResult.HasValue && wpfStackedWaterfall2.DialogResult.Value)
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
			A = wpfStackedWaterfall2.Range;
			result.B = wpfStackedWaterfall2.chkPreserveFormulas.IsChecked.Value;
			result.A = (float)wpfStackedWaterfall2.numChartWidth.Value.Value;
			result.B = (float)wpfStackedWaterfall2.numChartHeight.Value.Value;
			if (wpfStackedWaterfall2.radLabelAbove.IsChecked == true)
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
				result.A = JD.B;
			}
			else if (wpfStackedWaterfall2.radLabelAboveBelow.IsChecked == true)
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
				result.A = JD.A;
			}
			else
			{
				result.A = JD.C;
			}
			System.Windows.Media.Color buttonColor = QuickCharts2.GetButtonColor(wpfStackedWaterfall2.btnBridgeLine);
			result.A = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(buttonColor.R, buttonColor.G, buttonColor.B));
			result.A = wpfStackedWaterfall2.chkErrorBars.IsChecked.Value;
			QuickCharts2.SetChartWidth(ref xmlDoc, new decimal(wpfStackedWaterfall2.numChartWidth.Value.Value));
			QuickCharts2.SetChartHeight(ref xmlDoc, new decimal(wpfStackedWaterfall2.numChartHeight.Value.Value));
			QuickCharts2.SetPreserveFormulas(ref xmlDoc, wpfStackedWaterfall2.chkPreserveFormulas.IsChecked.Value);
			xmlDoc.SelectSingleNode(StackedWaterfall.m_A + VH.A(60864)).InnerText = (0 - (wpfStackedWaterfall2.chkErrorBars.IsChecked.Value ? 1 : 0)).ToString();
			result.A = xmlDoc.SelectSingleNode(StackedWaterfall.m_A + VH.A(60895)).InnerText;
			XmlNode xmlNode = xmlDoc.SelectSingleNode(StackedWaterfall.m_A + VH.A(55388));
			int a = (int)result.A;
			xmlNode.InnerText = a.ToString();
			xmlDoc.SelectSingleNode(StackedWaterfall.m_A + VH.A(60841)).InnerText = QuickCharts2.Color2RGB(buttonColor);
			KH.A.SaveSettings(xmlDoc);
		}
		else
		{
			B = false;
		}
		wpfStackedWaterfall2 = null;
		xmlDoc = null;
		return result;
	}

	private static void A(LD A, Worksheet B, ref Range C, ref Range D, int E, int F, int G)
	{
		Microsoft.Office.Interop.Excel.Application application = C.Application;
		int num = Conversions.ToInteger(C.Rows.CountLarge);
		int num2 = Conversions.ToInteger(C.Columns.CountLarge);
		checked
		{
			if (A.B)
			{
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				((Range)B.Cells[E, 1]).PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				int num3 = num;
				for (int i = 1; i <= num3; i++)
				{
					int num4 = num2;
					for (int j = 1; j <= num4; j++)
					{
						Range range = (Range)C.Cells[i, j];
						Range range2 = (Range)B.Cells[E - 1 + i, j];
						if (Operators.ConditionalCompareObjectEqual(range.Formula, string.Empty, TextCompare: false))
						{
							range2.Clear();
						}
						else if (Conversions.ToBoolean(range.HasFormula))
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
							string numberFormat = Conversions.ToString(range.NumberFormat);
							string formula = range.Formula.ToString();
							Range range3 = ((_Worksheet)C.Worksheet).get_Range((object)VH.A(60932), RuntimeHelpers.GetObjectValue(Missing.Value));
							range3.Formula = formula;
							range3.Cut(range2);
							_ = null;
							Range obj = (Range)B.Cells[E - 1 + i, j];
							obj.Font.Color = G;
							obj.NumberFormat = numberFormat;
							_ = null;
						}
						else if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.Value2)) || KH.A.AutoColorText)
						{
							range2.Font.Color = F;
						}
						range = null;
						range2 = null;
					}
					while (true)
					{
						switch (5)
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
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
				C = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 1]), RuntimeHelpers.GetObjectValue(B.Cells[Operators.SubtractObject(Operators.AddObject(E, C.Rows.CountLarge), 1), RuntimeHelpers.GetObjectValue(C.Columns.CountLarge)]));
			}
			else
			{
				Range range4 = C;
				C.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
				Range obj2 = (Range)B.Cells[E, 1];
				obj2.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				obj2.Select();
				B.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), true);
				D = ((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E + 1, 2]), RuntimeHelpers.GetObjectValue(B.Cells[E + num - 1, num2]));
				((_Worksheet)B).get_Range(RuntimeHelpers.GetObjectValue(B.Cells[E, 1]), RuntimeHelpers.GetObjectValue(B.Cells[E + num, num2])).Font.Color = G;
				C = (Range)application.Selection;
				int num5 = Conversions.ToInteger(range4.Rows.CountLarge);
				for (int k = 1; k <= num5; k++)
				{
					int num6 = Conversions.ToInteger(range4.Columns.CountLarge);
					for (int l = 1; l <= num6; l++)
					{
						if (!Operators.ConditionalCompareObjectEqual(((Range)range4.Cells[k, l]).Formula, string.Empty, TextCompare: false))
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
						((Range)C.Cells[k, l]).Value2 = "";
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0522;
						}
						continue;
						end_IL_0522:
						break;
					}
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
			application.CutCopyMode = (XlCutCopyMode)0;
			application = null;
		}
	}

	private static MD A(Worksheet A, ChartObject B, LD C)
	{
		int column = B.BottomRightCell.Column;
		int num = 4;
		MD result = default(MD);
		checked
		{
			int num2 = column + 1;
			Range range = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[2, num2]), RuntimeHelpers.GetObjectValue(A.Cells[2, column + 2]));
			range.Interior.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			range.VerticalAlignment = XlVAlign.xlVAlignCenter;
			range.RowHeight = 22;
			Range obj = (Range)range.Cells[1, 1];
			obj.Value2 = VH.A(60947);
			obj.Font.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
			obj.Font.Size = 14;
			_ = null;
			_ = null;
			((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[num - 1, num2]), RuntimeHelpers.GetObjectValue(A.Cells[StackedWaterfall.m_A, column + 2])).Interior.Color = QuickCharts2.OPTIONS_TABLE_FILL;
			Range range2 = (Range)A.Cells[num - 1, num2];
			Range range3 = range2;
			range3.Value2 = true;
			try
			{
				range3.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			CheckBox checkBox = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
			{
				range3.Left,
				range3.Top,
				range3.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)2).Width,
				range3.Height
			}, null, null, null);
			checkBox.Text = VH.A(60976);
			checkBox.LinkedCell = range3.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			if (C.A)
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
				checkBox.Value = Microsoft.Office.Interop.Excel.Constants.xlBoth;
			}
			else
			{
				checkBox.Value = Microsoft.Office.Interop.Excel.Constants.xlOff;
			}
			range3 = null;
			result.A = range2;
			range2 = (Range)A.Cells[num, num2];
			Range range4 = range2;
			range4.Value2 = true;
			try
			{
				range4.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			checkBox = (CheckBox)NewLateBinding.LateGet(A.CheckBoxes(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(60813), new object[4]
			{
				range4.Left,
				range4.Top,
				range4.get_Resize(RuntimeHelpers.GetObjectValue(Missing.Value), (object)2).Width,
				range4.Height
			}, null, null, null);
			checkBox.Text = VH.A(61011);
			checkBox.LinkedCell = range4.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			checkBox.Value = Microsoft.Office.Interop.Excel.Constants.xlBoth;
			range4 = null;
			result.B = range2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num + 2, num2], VH.A(61046));
			QuickCharts2.FormatOptionsInput(range2.get_Offset((object)3, (object)0));
			result.C = range2.get_Offset((object)3, (object)0);
			Range c = result.C;
			c.get_Address((object)true, (object)true, XlReferenceStyle.xlR1C1, (object)false, RuntimeHelpers.GetObjectValue(Missing.Value));
			c.Value2 = 0;
			try
			{
				c.NumberFormat = VH.A(20593);
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			c = null;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num + 5, num2], VH.A(61097));
			range2 = (Range)A.Cells[num + 6, num2];
			Range range5 = range2;
			DropDown dropDown = (DropDown)NewLateBinding.LateGet(A.DropDowns(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(61140), new object[4] { range5.Left, range5.Top, range5.Width, range5.Height }, null, null, null);
			dropDown.AddItem(VH.A(61147), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.AddItem(VH.A(56560), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.AddItem(VH.A(61174), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.LinkedCell = range5.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			JD a = C.A;
			if (a != JD.A)
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
				if (a != JD.B)
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
					dropDown.ListIndex = 3;
				}
				else
				{
					dropDown.ListIndex = 2;
				}
			}
			else
			{
				dropDown.ListIndex = 1;
			}
			try
			{
				range5.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			range5 = null;
			result.D = range2;
			QuickCharts2.FormatOptionsHeader((Range)A.Cells[num + 8, num2], VH.A(61183));
			range2 = (Range)A.Cells[num + 9, num2];
			Range range6 = range2;
			dropDown = (DropDown)NewLateBinding.LateGet(A.DropDowns(RuntimeHelpers.GetObjectValue(Missing.Value)), null, VH.A(61140), new object[4] { range6.Left, range6.Top, range6.Width, range6.Height }, null, null, null);
			dropDown.AddItem(VH.A(61226), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.AddItem(VH.A(61263), RuntimeHelpers.GetObjectValue(Missing.Value));
			dropDown.LinkedCell = range6.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			if (Operators.CompareString(C.A, Conversions.ToString(1), TextCompare: false) == 0)
			{
				dropDown.ListIndex = 1;
			}
			else
			{
				dropDown.ListIndex = 2;
			}
			try
			{
				range6.NumberFormat = QuickCharts2.NUMFORMAT_HIDDEN;
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
			range6 = null;
			Range range7 = range2.get_Offset((object)1, RuntimeHelpers.GetObjectValue(Missing.Value));
			range7.FormulaR1C1 = VH.A(61300);
			range7.Font.Italic = true;
			range7.Font.Color = Information.RGB(192, 0, 0);
			_ = null;
			range2.get_Offset((object)2, RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Color = QuickCharts2.OPTIONS_TABLE_FILL;
			_ = null;
			result.E = range2.get_Offset((object)2, RuntimeHelpers.GetObjectValue(Missing.Value));
			Border border = ((_Worksheet)A).get_Range(RuntimeHelpers.GetObjectValue(A.Cells[StackedWaterfall.m_A, column + 1]), RuntimeHelpers.GetObjectValue(A.Cells[StackedWaterfall.m_A, column + 2])).Borders[XlBordersIndex.xlEdgeBottom];
			border.LineStyle = XlLineStyle.xlContinuous;
			border.Weight = 2;
			border.Color = QuickCharts2.OPTIONS_DARK_YELLOW;
			_ = null;
			_ = null;
			checkBox = null;
			range2 = null;
			return result;
		}
	}

	public static Range Example(Worksheet ws)
	{
		Worksheet worksheet = ws;
		((_Worksheet)worksheet).get_Range((object)VH.A(61417), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61422), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57265);
		((_Worksheet)worksheet).get_Range((object)VH.A(61427), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61432);
		((_Worksheet)worksheet).get_Range((object)VH.A(61439), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61444);
		((_Worksheet)worksheet).get_Range((object)VH.A(61451), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61456);
		((_Worksheet)worksheet).get_Range((object)VH.A(61467), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61472);
		((_Worksheet)worksheet).get_Range((object)VH.A(61481), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57276);
		((_Worksheet)worksheet).get_Range((object)VH.A(61486), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61491);
		((_Worksheet)worksheet).get_Range((object)VH.A(61502), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 150;
		((_Worksheet)worksheet).get_Range((object)VH.A(61507), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61512), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -25;
		((_Worksheet)worksheet).get_Range((object)VH.A(61517), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61522), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61527), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61532), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61537);
		((_Worksheet)worksheet).get_Range((object)VH.A(61546), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 125;
		((_Worksheet)worksheet).get_Range((object)VH.A(61551), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61556), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61561), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61566), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61571), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61576), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61581);
		((_Worksheet)worksheet).get_Range((object)VH.A(61592), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 100;
		((_Worksheet)worksheet).get_Range((object)VH.A(61597), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 25;
		((_Worksheet)worksheet).get_Range((object)VH.A(61602), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61607), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61612), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61617), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		Range range = ((_Worksheet)worksheet).get_Range((object)VH.A(61622), RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			range.NumberFormat = QuickCharts2.CURRENCY_FORMAT_1;
			range.Font.Color = QuickCharts2.InputColor();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		range = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61633), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61638), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57265);
		((_Worksheet)worksheet).get_Range((object)VH.A(61643), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61432);
		((_Worksheet)worksheet).get_Range((object)VH.A(61648), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61444);
		((_Worksheet)worksheet).get_Range((object)VH.A(61653), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61456);
		((_Worksheet)worksheet).get_Range((object)VH.A(61658), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61472);
		((_Worksheet)worksheet).get_Range((object)VH.A(61663), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57276);
		((_Worksheet)worksheet).get_Range((object)VH.A(61668), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61491);
		((_Worksheet)worksheet).get_Range((object)VH.A(61673), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61678), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61683), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -25;
		((_Worksheet)worksheet).get_Range((object)VH.A(61688), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61693), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61698), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61703), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61537);
		((_Worksheet)worksheet).get_Range((object)VH.A(61708), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61713), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61718), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61723), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61728), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61733), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61738), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61581);
		((_Worksheet)worksheet).get_Range((object)VH.A(61743), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61748), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 25;
		((_Worksheet)worksheet).get_Range((object)VH.A(61753), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61758), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61763), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61768), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61773), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57265);
		((_Worksheet)worksheet).get_Range((object)VH.A(61778), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 375;
		((_Worksheet)worksheet).get_Range((object)VH.A(61783), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61788), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61793), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61798), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61803), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		Range range2 = ((_Worksheet)worksheet).get_Range((object)VH.A(61808), RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			range2.NumberFormat = QuickCharts2.CURRENCY_FORMAT_1;
			range2.Font.Color = QuickCharts2.InputColor();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		range2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61819), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61826), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57265);
		((_Worksheet)worksheet).get_Range((object)VH.A(61833), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61432);
		((_Worksheet)worksheet).get_Range((object)VH.A(61840), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61444);
		((_Worksheet)worksheet).get_Range((object)VH.A(61847), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61854);
		((_Worksheet)worksheet).get_Range((object)VH.A(61867), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61456);
		((_Worksheet)worksheet).get_Range((object)VH.A(61874), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61472);
		((_Worksheet)worksheet).get_Range((object)VH.A(61881), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57276);
		((_Worksheet)worksheet).get_Range((object)VH.A(61888), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61491);
		((_Worksheet)worksheet).get_Range((object)VH.A(61895), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 150;
		((_Worksheet)worksheet).get_Range((object)VH.A(61902), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61909), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -25;
		((_Worksheet)worksheet).get_Range((object)VH.A(61916), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61923), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61930), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61937), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61944), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61537);
		((_Worksheet)worksheet).get_Range((object)VH.A(61951), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 125;
		((_Worksheet)worksheet).get_Range((object)VH.A(61958), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61965), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61972), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(61979), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(61986), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(61993), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62000), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61581);
		((_Worksheet)worksheet).get_Range((object)VH.A(62007), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 100;
		((_Worksheet)worksheet).get_Range((object)VH.A(62014), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 25;
		((_Worksheet)worksheet).get_Range((object)VH.A(62021), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(62028), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62035), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62042), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62049), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		Range range3 = ((_Worksheet)worksheet).get_Range((object)VH.A(62056), RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			range3.NumberFormat = QuickCharts2.CURRENCY_FORMAT_1;
			range3.Font.Color = QuickCharts2.InputColor();
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		range3 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62071), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62078), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57265);
		((_Worksheet)worksheet).get_Range((object)VH.A(62085), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61432);
		((_Worksheet)worksheet).get_Range((object)VH.A(62092), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61444);
		((_Worksheet)worksheet).get_Range((object)VH.A(62099), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61854);
		((_Worksheet)worksheet).get_Range((object)VH.A(62106), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61456);
		((_Worksheet)worksheet).get_Range((object)VH.A(62113), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61472);
		((_Worksheet)worksheet).get_Range((object)VH.A(62120), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57276);
		((_Worksheet)worksheet).get_Range((object)VH.A(62127), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61491);
		((_Worksheet)worksheet).get_Range((object)VH.A(62134), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62141), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 75;
		((_Worksheet)worksheet).get_Range((object)VH.A(62148), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -25;
		((_Worksheet)worksheet).get_Range((object)VH.A(62155), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62162), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62169), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(62176), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62183), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61537);
		((_Worksheet)worksheet).get_Range((object)VH.A(62190), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62197), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62204), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62211), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62218), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(62225), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62232), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62239), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(61581);
		((_Worksheet)worksheet).get_Range((object)VH.A(62246), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62253), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 25;
		((_Worksheet)worksheet).get_Range((object)VH.A(62260), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = -75;
		((_Worksheet)worksheet).get_Range((object)VH.A(62267), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62274), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62281), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 50;
		((_Worksheet)worksheet).get_Range((object)VH.A(62288), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62295), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = VH.A(57265);
		((_Worksheet)worksheet).get_Range((object)VH.A(62302), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = 375;
		((_Worksheet)worksheet).get_Range((object)VH.A(62309), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62316), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62323), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62330), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62337), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		((_Worksheet)worksheet).get_Range((object)VH.A(62344), RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = null;
		Range range4 = ((_Worksheet)worksheet).get_Range((object)VH.A(62351), RuntimeHelpers.GetObjectValue(Missing.Value));
		try
		{
			range4.NumberFormat = QuickCharts2.CURRENCY_FORMAT_1;
			range4.Font.Color = QuickCharts2.InputColor();
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		range4 = null;
		Range result = ((_Worksheet)worksheet).get_Range((object)VH.A(62366), RuntimeHelpers.GetObjectValue(Missing.Value));
		worksheet = null;
		return result;
	}

	private static object A(Range A, bool B = false, bool C = false, long D = 0L, long E = 0L, bool F = true, bool G = true)
	{
		Range range = A.Areas[1].EntireRow;
		Range range2 = A.Areas[1].EntireColumn;
		checked
		{
			if (A.Areas.Count > 1)
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
				int count = A.Areas.Count;
				for (int i = 2; i <= count; i++)
				{
					range = MH.A.Application.Union(range, A.Areas[i].EntireRow, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					range2 = MH.A.Application.Union(range2, A.Areas[i].EntireColumn, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
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
			List<long> list = new List<long>();
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.Areas.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Range range3 = (Range)enumerator.Current;
					{
						enumerator2 = range3.Rows.GetEnumerator();
						try
						{
							while (enumerator2.MoveNext())
							{
								Range range4 = (Range)enumerator2.Current;
								list.Add(range4.Row);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_03b4;
								}
								continue;
								end_IL_03b4:
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
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_03e1;
					}
					continue;
					end_IL_03e1:
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
			list.Sort();
			List<long> list2 = new List<long>();
			IEnumerator enumerator3 = default(IEnumerator);
			try
			{
				enumerator3 = range2.Areas.GetEnumerator();
				IEnumerator enumerator4 = default(IEnumerator);
				while (enumerator3.MoveNext())
				{
					Range range5 = (Range)enumerator3.Current;
					try
					{
						enumerator4 = range5.Columns.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							Range range6 = (Range)enumerator4.Current;
							list2.Add(range6.Column);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0476;
							}
							continue;
							end_IL_0476:
							break;
						}
					}
					finally
					{
						if (enumerator4 is IDisposable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator4 as IDisposable).Dispose();
								break;
							}
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
						goto end_IL_04b0;
					}
					continue;
					end_IL_04b0:
					break;
				}
			}
			finally
			{
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator3 as IDisposable).Dispose();
						break;
					}
				}
			}
			list2.Sort();
			if (C)
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
				if (D < 0)
				{
					D = list.Max();
				}
				if (E < 0)
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
					E = list2.Max();
				}
			}
			object parent = A.Parent;
			object[,] array;
			if (B)
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
				array = new object[(int)list2.Max() + 1, (int)list.Max() + 1];
				int num = (int)(list.Max() - list.Min());
				for (int j = 0; j <= num; j++)
				{
					int num2 = (int)list2.Min();
					int num3 = (int)list2.Max();
					for (int k = num2; k <= num3; k++)
					{
						string text = VH.A(62377);
						if ((list[j] <= D) | (list2[k] <= E))
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
							text = VH.A(62386);
						}
						if (C)
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
							object instance = parent;
							string memberName = VH.A(62391);
							List<long> list3;
							int index;
							List<long> list4;
							int index2;
							object[] obj = new object[2]
							{
								(list3 = list)[index = j],
								(list4 = list2)[index2 = k]
							};
							object[] array2 = obj;
							bool[] obj2 = new bool[2] { true, true };
							bool[] array3 = obj2;
							object instance2 = NewLateBinding.LateGet(instance, null, memberName, obj, null, null, obj2);
							if (array3[0])
							{
								list3[index] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(long));
							}
							if (array3[1])
							{
								list4[index2] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(long));
							}
							object[] array4;
							bool[] array5;
							object obj3 = NewLateBinding.LateGet(instance2, null, VH.A(5814), array4 = new object[4]
							{
								F,
								G,
								XlReferenceStyle.xlA1,
								true
							}, null, null, array5 = new bool[4] { true, true, false, false });
							if (array5[0])
							{
								F = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[0]), typeof(bool));
							}
							if (array5[1])
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
								G = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[1]), typeof(bool));
							}
							string text2 = Conversions.ToString(obj3);
							array[k, j] = string.Format(VH.A(62402), text2, text2, text, text);
						}
						else
						{
							object[,] array6 = array;
							int num4 = k;
							int num5 = j;
							string format = VH.A(48804);
							object instance3 = parent;
							string memberName2 = VH.A(62391);
							List<long> list4;
							int index2;
							List<long> list3;
							int index;
							object[] obj4 = new object[2]
							{
								(list4 = list)[index2 = j],
								(list3 = list2)[index = k]
							};
							object[] array2 = obj4;
							bool[] obj5 = new bool[2] { true, true };
							bool[] array3 = obj5;
							object instance4 = NewLateBinding.LateGet(instance3, null, memberName2, obj4, null, null, obj5);
							if (array3[0])
							{
								list4[index2] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(long));
							}
							if (array3[1])
							{
								list3[index] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(long));
							}
							object[] array4;
							bool[] array5;
							object obj6 = NewLateBinding.LateGet(instance4, null, VH.A(5814), array4 = new object[4]
							{
								F,
								G,
								XlReferenceStyle.xlA1,
								true
							}, null, null, array5 = new bool[4] { true, true, false, false });
							if (array5[0])
							{
								F = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[0]), typeof(bool));
							}
							if (array5[1])
							{
								G = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[1]), typeof(bool));
							}
							array6[num4, num5] = string.Format(format, RuntimeHelpers.GetObjectValue(obj6));
						}
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0948;
						}
						continue;
						end_IL_0948:
						break;
					}
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
			else
			{
				array = new object[(int)(list.Max() - list.Min()) + 1, (int)(list2.Max() - list2.Min()) + 1];
				int index = (int)(list.Max() - list.Min());
				for (int l = 0; l <= index; l++)
				{
					int index2 = (int)(list2.Max() - list2.Min());
					for (int m = 0; m <= index2; m++)
					{
						string text3 = VH.A(62377);
						if (unchecked(l <= D || m <= E))
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
							text3 = VH.A(62386);
						}
						List<long> list4;
						int index4;
						List<long> list3;
						int index3;
						object[] array2;
						bool[] array3;
						object[] array4;
						bool[] array5;
						if (C)
						{
							object instance5 = parent;
							string memberName3 = VH.A(62391);
							object[] obj7 = new object[2]
							{
								(list3 = list)[index3 = l],
								(list4 = list2)[index4 = m]
							};
							array2 = obj7;
							bool[] obj8 = new bool[2] { true, true };
							array3 = obj8;
							object instance6 = NewLateBinding.LateGet(instance5, null, memberName3, obj7, null, null, obj8);
							if (array3[0])
							{
								list3[index3] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(long));
							}
							if (array3[1])
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
								list4[index4] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(long));
							}
							object obj9 = NewLateBinding.LateGet(instance6, null, VH.A(5814), array4 = new object[4]
							{
								F,
								G,
								XlReferenceStyle.xlA1,
								true
							}, null, null, array5 = new bool[4] { true, true, false, false });
							if (array5[0])
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
								F = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[0]), typeof(bool));
							}
							if (array5[1])
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
								G = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[1]), typeof(bool));
							}
							string text4 = Conversions.ToString(obj9);
							array[l, m] = string.Format(VH.A(62402), text4, text4, text3, text3);
							continue;
						}
						object[,] array7 = array;
						int num6 = l;
						int num7 = m;
						string format2 = VH.A(48804);
						object instance7 = parent;
						string memberName4 = VH.A(62391);
						object[] obj10 = new object[2]
						{
							(list4 = list)[index4 = l],
							(list3 = list2)[index3 = m]
						};
						array2 = obj10;
						bool[] obj11 = new bool[2] { true, true };
						array3 = obj11;
						object instance8 = NewLateBinding.LateGet(instance7, null, memberName4, obj10, null, null, obj11);
						if (array3[0])
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
							list4[index4] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[0]), typeof(long));
						}
						if (array3[1])
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
							list3[index3] = (long)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array2[1]), typeof(long));
						}
						object obj12 = NewLateBinding.LateGet(instance8, null, VH.A(5814), array4 = new object[4]
						{
							F,
							G,
							XlReferenceStyle.xlA1,
							true
						}, null, null, array5 = new bool[4] { true, true, false, false });
						if (array5[0])
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
							F = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[0]), typeof(bool));
						}
						if (array5[1])
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
							G = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array4[1]), typeof(bool));
						}
						array7[num6, num7] = string.Format(format2, RuntimeHelpers.GetObjectValue(obj12));
					}
				}
			}
			parent = null;
			return array;
		}
	}

	private static void A(Range A)
	{
		bool screenUpdating = MH.A.Application.ScreenUpdating;
		MH.A.Application.ScreenUpdating = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Areas.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				XlBordersIndex xlBordersIndex = XlBordersIndex.xlDiagonalDown;
				do
				{
					switch (xlBordersIndex)
					{
					case XlBordersIndex.xlDiagonalDown:
					case XlBordersIndex.xlDiagonalUp:
						range.Borders[xlBordersIndex].LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
						break;
					case XlBordersIndex.xlEdgeLeft:
					case XlBordersIndex.xlEdgeTop:
					case XlBordersIndex.xlEdgeBottom:
					case XlBordersIndex.xlEdgeRight:
					{
						Border border2 = range.Borders[xlBordersIndex];
						border2.Color = 9671571L;
						border2.LineStyle = XlLineStyle.xlContinuous;
						border2.Weight = XlBorderWeight.xlThin;
						_ = null;
						break;
					}
					case XlBordersIndex.xlInsideVertical:
					case XlBordersIndex.xlInsideHorizontal:
					{
						Border border = range.Borders[xlBordersIndex];
						border.Color = 14277081L;
						border.LineStyle = XlLineStyle.xlContinuous;
						border.Weight = XlBorderWeight.xlThin;
						_ = null;
						break;
					}
					}
					xlBordersIndex++;
				}
				while (xlBordersIndex <= XlBordersIndex.xlInsideHorizontal);
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
					range = null;
					break;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0135;
				}
				continue;
				end_IL_0135:
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
		MH.A.Application.ScreenUpdating = screenUpdating;
	}

	private static void B(Range A)
	{
		bool screenUpdating = MH.A.Application.ScreenUpdating;
		MH.A.Application.ScreenUpdating = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Areas.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				XlBordersIndex xlBordersIndex = XlBordersIndex.xlDiagonalDown;
				do
				{
					switch (xlBordersIndex)
					{
					case XlBordersIndex.xlDiagonalDown:
					case XlBordersIndex.xlDiagonalUp:
						range.Borders[xlBordersIndex].LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
						break;
					case XlBordersIndex.xlEdgeLeft:
					case XlBordersIndex.xlEdgeTop:
					case XlBordersIndex.xlEdgeBottom:
					case XlBordersIndex.xlEdgeRight:
					{
						Border border = range.Borders[xlBordersIndex];
						border.Color = 14277081L;
						border.LineStyle = XlLineStyle.xlContinuous;
						border.Weight = XlBorderWeight.xlThin;
						_ = null;
						break;
					}
					case XlBordersIndex.xlInsideVertical:
					case XlBordersIndex.xlInsideHorizontal:
						range.Borders[xlBordersIndex].LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone;
						break;
					}
					xlBordersIndex++;
				}
				while (xlBordersIndex <= XlBordersIndex.xlInsideHorizontal);
				range = null;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		MH.A.Application.ScreenUpdating = screenUpdating;
	}

	private static MsoThemeColorIndex A(long A)
	{
		long num;
		checked
		{
			num = unchecked(checked(A - 1) % 6) + 1;
		}
		long num2 = num - 1;
		if ((ulong)num2 > 5uL)
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
		}
		else
		{
			switch (num2)
			{
			case 0L:
				return MsoThemeColorIndex.msoThemeColorAccent1;
			case 1L:
				return MsoThemeColorIndex.msoThemeColorAccent2;
			case 2L:
				return MsoThemeColorIndex.msoThemeColorAccent3;
			case 3L:
				return MsoThemeColorIndex.msoThemeColorAccent4;
			case 4L:
				return MsoThemeColorIndex.msoThemeColorAccent5;
			case 5L:
				return MsoThemeColorIndex.msoThemeColorAccent6;
			}
		}
		MsoThemeColorIndex result = default(MsoThemeColorIndex);
		return result;
	}

	private static double A(long A)
	{
		double num = Conversion.Int((double)checked(A - 1) / 6.0);
		if (num == 0.0)
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
					return 0.0;
				}
			}
		}
		if (num == 1.0)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return 0.15;
				}
			}
		}
		if (num == 2.0)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return 0.25;
				}
			}
		}
		if (num >= 3.0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return 0.35;
				}
			}
		}
		double result = default(double);
		return result;
	}

	private static void A(Series A, Range B)
	{
		A.HasDataLabels = true;
		A.HasLeaderLines = false;
		object instance = A.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
		NewLateBinding.LateSetComplex(instance, null, VH.A(62489), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateCall(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(60221), new object[0], null, null, null), null, VH.A(60595), new object[0], null, null, null), null, VH.A(60616), new object[0], null, null, null), null, VH.A(62508), new object[2]
		{
			7L,
			VH.A(48936) + B.get_Address((object)true, (object)true, MH.A.Application.ReferenceStyle, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value))
		}, null, null, null, IgnoreReturn: true);
		NewLateBinding.LateSetComplex(instance, null, VH.A(62541), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
		instance = null;
		_ = null;
	}

	private static string A(Range A)
	{
		string text = A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlR1C1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		return checked(Strings.Left(text, Strings.InStr(text, VH.A(7120)) - 1) + Strings.Mid(text, Strings.InStr(text, VH.A(43340)) + 1));
	}
}
