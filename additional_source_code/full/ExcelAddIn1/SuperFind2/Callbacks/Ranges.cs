using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.SuperFind2.Results;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Ranges
{
	internal static void A(WorksheetItem A, Microsoft.Office.Interop.Excel.Worksheet B)
	{
		A.A();
	}

	internal static void B(WorksheetItem A, Microsoft.Office.Interop.Excel.Worksheet B)
	{
		Microsoft.Office.Interop.Excel.Worksheet worksheet = B;
		if (Operators.CompareString(worksheet.PageSetup.PrintArea, "", TextCompare: false) != 0)
		{
			string[] array = Strings.Split(worksheet.PageSetup.PrintArea, CultureInfo.CurrentCulture.TextInfo.ListSeparator, -1, CompareMethod.Text);
			foreach (string cell in array)
			{
				A.G(((_Worksheet)worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)));
			}
		}
		worksheet = null;
	}

	internal static void A(WorksheetItem A, object B)
	{
		checked
		{
			Range range;
			Application application;
			if (B is Microsoft.Office.Interop.Excel.Worksheet)
			{
				Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)B;
				application = worksheet.Application;
				Watches watches = application.Watches;
				int num = watches.Count - 1;
				for (int i = 0; i <= num; i++)
				{
					range = (Range)watches.get_Item((object)i).Source;
					if (range.Worksheet != worksheet)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A.A(watches.get_Item((object)i));
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
				watches = null;
			}
			else
			{
				Microsoft.Office.Interop.Excel.Worksheet worksheet = ((Range)B).Worksheet;
				application = worksheet.Application;
				Watches watches2 = application.Watches;
				int num2 = watches2.Count - 1;
				for (int j = 0; j <= num2; j++)
				{
					range = (Range)watches2.get_Item((object)j).Source;
					if (range.Worksheet != worksheet)
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
					if (application.Intersect(range, (Range)B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
					{
						A.A(watches2.get_Item((object)j));
					}
				}
				watches2 = null;
			}
			range = null;
			application = null;
		}
	}

	internal static void A(WorksheetItem A, Range B)
	{
		Application application = B.Application;
		Microsoft.Office.Interop.Excel.Worksheet worksheet = B.Worksheet;
		Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent;
		IEnumerator enumerator = default(IEnumerator);
		if (workbook.Names.Count <= 100)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					{
						enumerator = workbook.Names.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Name name = (Name)enumerator.Current;
								Range range = null;
								try
								{
									range = name.RefersToRange;
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
								if (range != null)
								{
									if (range.Worksheet == worksheet)
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
										if (name.Visible)
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
											if (!ExcelAddIn1.Formulas.Names.A(name.Name))
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
													if (application.Intersect(range, B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
													{
														while (true)
														{
															switch (4)
															{
															case 0:
																break;
															default:
																A.E(range);
																goto end_IL_022f;
															}
															continue;
															end_IL_022f:
															break;
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
										}
									}
									range = null;
								}
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0261;
								}
								continue;
								end_IL_0261:
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
					application = null;
					worksheet = null;
					workbook = null;
					return;
				}
			}
		}
		Forms.WarningMessage(VH.A(102719));
		throw new Exception(VH.A(102911));
	}

	internal static void B(WorksheetItem A, object B)
	{
		int num = 0;
		Range range;
		Microsoft.Office.Interop.Excel.Worksheet worksheet;
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
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
			range = null;
			worksheet = (Microsoft.Office.Interop.Excel.Worksheet)B;
		}
		else
		{
			range = (Range)B;
			worksheet = range.Worksheet;
		}
		checked
		{
			foreach (Name item in Names.A((Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent))
			{
				if (item.Visible || num <= 25)
				{
					Ranges.A(A, item, range, D: true, E: false);
				}
				if (item.Visible)
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
				num++;
			}
			if (num > 25)
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
				A.C(num - 25 + VH.A(102942));
			}
			range = null;
			worksheet = null;
		}
	}

	internal static void C(WorksheetItem A, object B)
	{
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					using List<Name>.Enumerator enumerator = Names.A((Microsoft.Office.Interop.Excel.Workbook)((Microsoft.Office.Interop.Excel.Worksheet)B).Parent).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Name current = enumerator.Current;
						if (current.Visible)
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
							Ranges.A(A, current, null, D: true, E: false);
						}
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				}
			}
		}
		using List<Name>.Enumerator enumerator2 = Names.A((Microsoft.Office.Interop.Excel.Workbook)((Range)B).Worksheet.Parent).GetEnumerator();
		while (enumerator2.MoveNext())
		{
			Name current2 = enumerator2.Current;
			if (!current2.Visible)
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
			Ranges.A(A, current2, (Range)B, D: true, E: false);
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	internal static void D(WorksheetItem A, object B)
	{
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			while (true)
			{
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
					using List<Name>.Enumerator enumerator = Names.A((Microsoft.Office.Interop.Excel.Workbook)((Microsoft.Office.Interop.Excel.Worksheet)B).Parent).GetEnumerator();
					while (enumerator.MoveNext())
					{
						Name current = enumerator.Current;
						Ranges.A(A, current, null, D: false, E: true);
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
			}
		}
		using List<Name>.Enumerator enumerator2 = Names.A((Microsoft.Office.Interop.Excel.Workbook)((Range)B).Worksheet.Parent).GetEnumerator();
		while (enumerator2.MoveNext())
		{
			Name current2 = enumerator2.Current;
			Ranges.A(A, current2, (Range)B, D: false, E: true);
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private static void A(WorksheetItem A, Name B, Range C, bool D, bool E)
	{
		Range range = null;
		try
		{
			range = B.RefersToRange;
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
				switch (6)
				{
				case 0:
					break;
				default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (range.Worksheet == A.Worksheet)
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
							if (C != null)
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
								if (B.Application.Intersect(range, C, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
								{
									goto IL_0206;
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
							if (E)
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
								if (Names.A(B))
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
									A.P(range);
									goto IL_0206;
								}
							}
							if (D)
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
								A.A(B, range, C: false);
							}
						}
						goto IL_0206;
					}
					IL_0206:
					range = null;
					return;
				}
			}
		}
		if (!D)
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
			try
			{
				string text = B.RefersTo.ToString();
				if (!text.Contains(A.Worksheet.Name))
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
					if (Regex.IsMatch(text.Replace(A.Worksheet.Name, ""), VH.A(102981)))
					{
						A.A(B, null, C: true, VH.A(103000));
					}
					return;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Ranges.A(A, B, Ranges.A);
	}

	private static bool A(int A)
	{
		return A == 1;
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Ranges.A(A, B, Ranges.B);
	}

	private static bool B(int A)
	{
		return A == 0;
	}

	private static void A(WorksheetItem A, Range B, Func<int, bool> C)
	{
		try
		{
			B = (Range)B.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = B.Cells.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					if (C(range.Row % 2))
					{
						A.F(range.EntireRow);
					}
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
					return;
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Ranges.B(A, B, C);
	}

	private static bool C(int A)
	{
		return A == 1;
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Ranges.B(A, B, D);
	}

	private static bool D(int A)
	{
		return A == 0;
	}

	private static void B(WorksheetItem A, Range B, Func<int, bool> C)
	{
		try
		{
			B = (Range)B.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = B.Cells.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					if (!C(range.Column % 2))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A.F(range.EntireColumn);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return;
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
