using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Formulas;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Visualizations;

public sealed class FormulaReport
{
	public static void Create()
	{
		Application application = MH.A.Application;
		Dictionary<string, int> dictionary = new Dictionary<string, int>();
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		checked
		{
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook;
			Worksheet worksheet;
			Range range;
			try
			{
				activeWorkbook = application.ActiveWorkbook;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = activeWorkbook.Worksheets.GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					IEnumerator enumerator3 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						worksheet = (Worksheet)enumerator.Current;
						range = null;
						try
						{
							range = ExcelAddIn1.Formulas.Helpers.SpecialCellsFormulas(worksheet.UsedRange);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (range == null)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						try
						{
							enumerator2 = range.Cells.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								MatchCollection matchCollection = Regex.Matches(Conversions.ToString(((Range)enumerator2.Current).Formula), VH.A(51280));
								{
									enumerator3 = matchCollection.GetEnumerator();
									try
									{
										while (enumerator3.MoveNext())
										{
											string key = ((Match)enumerator3.Current).Groups[2].ToString();
											if (dictionary.ContainsKey(key))
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
												dictionary[key]++;
											}
											else
											{
												dictionary.Add(key, 1);
											}
										}
										while (true)
										{
											switch (4)
											{
											case 0:
												break;
											default:
												goto end_IL_0146;
											}
											continue;
											end_IL_0146:
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
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_0176;
								}
								continue;
								end_IL_0176:
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
					while (true)
					{
						switch (1)
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
				worksheet = (Worksheet)activeWorkbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(activeWorkbook.Sheets[1]), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				((Range)worksheet.Cells[1, 1]).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(51343));
				((Range)worksheet.Cells[3, 1]).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(51372));
				((Range)worksheet.Cells[3, 2]).set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)VH.A(51389));
				int num = 4;
				foreach (KeyValuePair<string, int> item in dictionary)
				{
					worksheet.Cells[num, 1] = item.Key;
					worksheet.Cells[num, 2] = item.Value;
					num++;
				}
				Range obj = (Range)worksheet.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
				obj.AutoFit();
				Range range2 = obj;
				obj.ColumnWidth = Operators.AddObject(range2.ColumnWidth, 0);
				_ = null;
				Range obj2 = (Range)worksheet.Columns[2, RuntimeHelpers.GetObjectValue(Missing.Value)];
				obj2.AutoFit();
				range2 = obj2;
				obj2.ColumnWidth = Operators.AddObject(range2.ColumnWidth, 0);
				_ = null;
				IEnumerator enumerator5 = default(IEnumerator);
				try
				{
					enumerator5 = ((_Worksheet)worksheet).get_Range((object)VH.A(51412), RuntimeHelpers.GetObjectValue(Missing.Value)).Cells.GetEnumerator();
					while (enumerator5.MoveNext())
					{
						Range obj3 = (Range)enumerator5.Current;
						obj3.AutoFilter(obj3.Column, RuntimeHelpers.GetObjectValue(Missing.Value), XlAutoFilterOperator.xlAnd, RuntimeHelpers.GetObjectValue(Missing.Value), true);
					}
				}
				finally
				{
					if (enumerator5 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator5 as IDisposable).Dispose();
							break;
						}
					}
				}
				_ = null;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application = null;
			activeWorkbook = null;
			worksheet = null;
			range = null;
			dictionary = null;
		}
	}
}
