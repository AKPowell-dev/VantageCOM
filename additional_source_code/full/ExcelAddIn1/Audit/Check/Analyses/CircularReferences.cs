using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class CircularReferences
{
	[CompilerGenerated]
	internal sealed class M
	{
		public Microsoft.Office.Interop.Excel.Workbook A;

		public string A;

		public Microsoft.Office.Interop.Excel.Workbook B;

		public Application A;

		public M(M A)
		{
			if (A == null)
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
				this.A = A.A;
				this.A = A.A;
				this.B = A.B;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.SaveCopyAs(this.A);
		}

		[SpecialName]
		internal void B()
		{
			this.B = this.A.Workbooks.Add(this.A);
		}

		[SpecialName]
		internal void C()
		{
			File.Delete(this.A);
		}
	}

	internal static void A(Analysis A, Severity B, Microsoft.Office.Interop.Excel.Workbook C)
	{
		M a = default(M);
		M CS_0024_003C_003E8__locals28 = new M(a);
		CS_0024_003C_003E8__locals28.A = C;
		if (B == Severity.Ignore)
		{
			return;
		}
		CS_0024_003C_003E8__locals28.A = CS_0024_003C_003E8__locals28.A.Application;
		CS_0024_003C_003E8__locals28.B = null;
		XlCalculation calculation = CS_0024_003C_003E8__locals28.A.Calculation;
		bool iteration = CS_0024_003C_003E8__locals28.A.Iteration;
		XlGenerateTableRefs generateTableRefs = CS_0024_003C_003E8__locals28.A.GenerateTableRefs;
		bool flag = false;
		CircularReferences.A(CS_0024_003C_003E8__locals28.A);
		IEnumerator enumerator = CS_0024_003C_003E8__locals28.A.Worksheets.GetEnumerator();
		try
		{
			while (true)
			{
				if (enumerator.MoveNext())
				{
					if (((Microsoft.Office.Interop.Excel.Worksheet)enumerator.Current).CircularReference == null)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						flag = true;
						break;
					}
					break;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_00b0;
					}
					continue;
					end_IL_00b0:
					break;
				}
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
		Dictionary<string, string> dictionary;
		if (flag)
		{
			int a2 = A.A;
			try
			{
				CS_0024_003C_003E8__locals28.A = Path.Combine(I.A.FileSystem.SpecialDirectories.Temp, string.Format(VH.A(2161), Path.GetExtension(CS_0024_003C_003E8__locals28.A.FullName)));
				A.A(VH.A(2184), [SpecialName] () =>
				{
					CS_0024_003C_003E8__locals28.A.SaveCopyAs(CS_0024_003C_003E8__locals28.A);
				});
				A.A(VH.A(2211), [SpecialName] () =>
				{
					CS_0024_003C_003E8__locals28.B = CS_0024_003C_003E8__locals28.A.Workbooks.Add(CS_0024_003C_003E8__locals28.A);
				});
				A.A(VH.A(2236), [SpecialName] () =>
				{
					File.Delete(CS_0024_003C_003E8__locals28.A);
				});
				dictionary = new Dictionary<string, string>();
				A.ActionStarted(VH.A(2277), CS_0024_003C_003E8__locals28.B.Worksheets.Count);
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = CS_0024_003C_003E8__locals28.B.Worksheets.GetEnumerator();
					while (true)
					{
						if (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)enumerator2.Current;
							if (A.ItemCancelled())
							{
								break;
							}
							while (worksheet.CircularReference != null)
							{
								Range circularReference = worksheet.CircularReference;
								string key = worksheet.Name + VH.A(2312) + circularReference.FormulaR1C1.ToString();
								if (dictionary.ContainsKey(key))
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
									dictionary[key] = CS_0024_003C_003E8__locals28.A.Union(((_Worksheet)worksheet).get_Range((object)dictionary[key], RuntimeHelpers.GetObjectValue(Missing.Value)), worksheet.CircularReference, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								else
								{
									dictionary.Add(key, circularReference.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
								}
								circularReference.Value2 = 1;
								circularReference = null;
							}
							continue;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0444;
							}
							continue;
							end_IL_0444:
							break;
						}
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
				A.ActionEnded();
				if (A.A())
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0480;
						}
						continue;
						end_IL_0480:
						break;
					}
				}
				else
				{
					A.ActionStarted(VH.A(2319), dictionary.Keys.LongCount());
					string[] array = dictionary.Keys.ToArray();
					int num = 0;
					IEnumerator enumerator3 = default(IEnumerator);
					while (true)
					{
						if (num < array.Length)
						{
							string text = array[num];
							if (A.ItemCancelled())
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
								string[] array2 = CircularReferences.A(text);
								string text2 = string.Empty;
								Microsoft.Office.Interop.Excel.Worksheet worksheet2 = QB.A(CS_0024_003C_003E8__locals28.B, array2[0]);
								try
								{
									enumerator3 = ((_Worksheet)worksheet2).get_Range((object)dictionary[text], RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
									while (enumerator3.MoveNext())
									{
										Range range = (Range)enumerator3.Current;
										range.FormulaR1C1 = array2[1];
										worksheet2.Calculate();
										if (range.Worksheet.CircularReference == null)
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
										range.Value2 = 1;
										text2 = text2 + VH.A(2378) + range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									}
								}
								finally
								{
									if (enumerator3 is IDisposable)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											(enumerator3 as IDisposable).Dispose();
											break;
										}
									}
								}
								if (Operators.CompareString(text2, string.Empty, TextCompare: false) == 0)
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
									dictionary.Remove(text);
								}
								else
								{
									dictionary[text] = CS_0024_003C_003E8__locals28.A.Union(((_Worksheet)worksheet2).get_Range((object)Strings.Mid(text2, 2, text2.Length), RuntimeHelpers.GetObjectValue(Missing.Value)), ((_Worksheet)worksheet2).get_Range((object)Strings.Mid(text2, 2, text2.Length), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								worksheet2 = null;
								num = checked(num + 1);
								break;
							}
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
					A.ActionEnded();
					if (!A.A())
					{
						string[] array3 = dictionary.Keys.ToArray();
						foreach (string text3 in array3)
						{
							Microsoft.Office.Interop.Excel.Worksheet worksheet3 = QB.A(CS_0024_003C_003E8__locals28.A, CircularReferences.A(text3)[0]);
							A.Observations.Add(new CircularReference(B, ((_Worksheet)worksheet3).get_Range((object)dictionary[text3], RuntimeHelpers.GetObjectValue(Missing.Value))));
							worksheet3 = null;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_089a;
							}
							continue;
							end_IL_089a:
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
			finally
			{
				A.A(a2);
				if (CS_0024_003C_003E8__locals28.B != null)
				{
					try
					{
						CS_0024_003C_003E8__locals28.B.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
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
		CircularReferences.A(CS_0024_003C_003E8__locals28.A, iteration, calculation, generateTableRefs);
		CS_0024_003C_003E8__locals28.A = null;
		CS_0024_003C_003E8__locals28.B = null;
		dictionary = null;
	}

	internal static void A(List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Application application = C.Application;
			XlCalculation calculation = application.Calculation;
			bool iteration = application.Iteration;
			XlGenerateTableRefs generateTableRefs = application.GenerateTableRefs;
			Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
			CircularReferences.A(application);
			Dictionary<string, string> dictionary;
			if (C.CircularReference != null)
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
					C.Copy(C, RuntimeHelpers.GetObjectValue(Missing.Value));
					worksheet = (Microsoft.Office.Interop.Excel.Worksheet)application.ActiveSheet;
					worksheet.Visible = XlSheetVisibility.xlSheetVisible;
					dictionary = new Dictionary<string, string>();
					while (worksheet.CircularReference != null)
					{
						Range circularReference = worksheet.CircularReference;
						string key = circularReference.FormulaR1C1.ToString();
						if (dictionary.ContainsKey(key))
						{
							dictionary[key] = application.Union(((_Worksheet)worksheet).get_Range((object)dictionary[key], RuntimeHelpers.GetObjectValue(Missing.Value)), worksheet.CircularReference, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						else
						{
							dictionary.Add(key, circularReference.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
						}
						circularReference.Value2 = 1;
						circularReference = null;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						string[] array = dictionary.Keys.ToArray();
						foreach (string text in array)
						{
							string text2 = string.Empty;
							Microsoft.Office.Interop.Excel.Worksheet worksheet2 = worksheet;
							{
								enumerator = ((_Worksheet)worksheet2).get_Range((object)dictionary[text], RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
								try
								{
									while (enumerator.MoveNext())
									{
										Range range = (Range)enumerator.Current;
										range.FormulaR1C1 = text;
										worksheet2.Calculate();
										if (range.Worksheet.CircularReference == null)
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
										range.Value2 = 1;
										text2 = text2 + VH.A(2378) + range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											goto end_IL_039b;
										}
										continue;
										end_IL_039b:
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
							if (Operators.CompareString(text2, string.Empty, TextCompare: false) == 0)
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
								dictionary.Remove(text);
							}
							else
							{
								dictionary[text] = application.Union(((_Worksheet)worksheet2).get_Range((object)Strings.Mid(text2, 2, text2.Length), RuntimeHelpers.GetObjectValue(Missing.Value)), ((_Worksheet)worksheet2).get_Range((object)Strings.Mid(text2, 2, text2.Length), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							worksheet2 = null;
						}
						string[] array2 = dictionary.Keys.ToArray();
						foreach (string key2 in array2)
						{
							A.Add(new CircularReference(B, ((_Worksheet)C).get_Range((object)dictionary[key2], RuntimeHelpers.GetObjectValue(Missing.Value))));
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0615;
							}
							continue;
							end_IL_0615:
							break;
						}
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				finally
				{
					if (worksheet != null)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							worksheet.Delete();
							break;
						}
					}
				}
			}
			CircularReferences.A(application, iteration, calculation, generateTableRefs);
			application = null;
			worksheet = null;
			dictionary = null;
			return;
		}
	}

	private static string[] A(string A)
	{
		return A.Split(new string[1] { VH.A(2312) }, StringSplitOptions.RemoveEmptyEntries);
	}

	private static void A(Application A)
	{
		A.EnableEvents = false;
		A.DisplayAlerts = false;
		A.Calculation = XlCalculation.xlCalculationSemiautomatic;
		A.Iteration = false;
		A.GenerateTableRefs = XlGenerateTableRefs.xlGenerateTableRefA1;
		_ = null;
	}

	private static void A(Application A, bool B, XlCalculation C, XlGenerateTableRefs D)
	{
		A.Iteration = B;
		A.Calculation = C;
		A.GenerateTableRefs = D;
		A.DisplayAlerts = true;
		A.EnableEvents = true;
		_ = null;
	}
}
