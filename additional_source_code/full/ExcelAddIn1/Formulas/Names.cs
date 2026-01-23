using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class Names
{
	private struct ZF
	{
		public string A;

		public string B;

		public string C;

		public bool A;

		public string D;

		public string E;

		public bool B;

		public bool C;

		public bool D;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Comparison<ZF> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(ZF A, ZF B)
		{
			return B.A.Length.CompareTo(A.A.Length);
		}
	}

	[CompilerGenerated]
	internal sealed class AG
	{
		public Range A;

		public AG(AG A)
		{
			if (A == null)
			{
				return;
			}
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(Range A)
		{
			return Operators.CompareString(A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), this.A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0;
		}
	}

	public static readonly string NAME_PATTERN = VH.A(155625);

	internal static bool A(Name A)
	{
		if (!LikeOperator.LikeString(A.RefersTo.ToString(), VH.A(153926), CompareMethod.Binary))
		{
			return LikeOperator.LikeString(A.RefersTo.ToString(), VH.A(153939), CompareMethod.Binary);
		}
		return true;
	}

	internal static bool A(string A)
	{
		return Names.IsNative(A);
	}

	internal static bool B(Name A)
	{
		return Names.IsLinked(A);
	}

	internal static bool C(Name A)
	{
		bool result;
		try
		{
			result = A.RefersTo.ToString().StartsWith(VH.A(153954));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static bool A(Name A, bool B)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		bool result;
		if (!Names.A(A.Name))
		{
			if (!Names.B(A))
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
				List<Range> list = new List<Range>();
				Range activeCell = default(Range);
				if (!B)
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
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					activeCell = application.ActiveCell;
				}
				Range refersToRange;
				try
				{
					refersToRange = A.RefersToRange;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					result = false;
					ProjectData.ClearProjectError();
					goto IL_0116;
				}
				list = GetDependents(refersToRange, A);
				bool flag = default(bool);
				using (List<Range>.Enumerator enumerator = list.GetEnumerator())
				{
					while (true)
					{
						if (enumerator.MoveNext())
						{
							if (!Names.A(enumerator.Current, A))
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
								flag = true;
								break;
							}
							break;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_00c7;
							}
							continue;
							end_IL_00c7:
							break;
						}
						break;
					}
				}
				if (!B)
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
					Names.A(activeCell);
					application.ScreenUpdating = true;
					application.EnableEvents = true;
				}
				activeCell = null;
				refersToRange = null;
				list = null;
				result = flag;
			}
			else
			{
				result = true;
			}
		}
		else
		{
			result = true;
		}
		goto IL_0116;
		IL_0116:
		return result;
	}

	internal static bool A(Range A, Name B)
	{
		return Regex.IsMatch(NewLateBinding.LateGet(A, null, VH.A(1998), new object[0], null, null, null).ToString(), VH.A(4544) + B.Name + VH.A(4544));
	}

	internal static bool A(Name A, long B)
	{
		bool result;
		try
		{
			result = Conversions.ToLong(A.RefersToRange.CountLarge) <= B;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		finally
		{
		}
		return result;
	}

	internal static bool D(Name A)
	{
		if (A.Name.EndsWith(VH.A(153971)))
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
					return !A.Visible;
				}
			}
		}
		return false;
	}

	internal static bool E(Name A)
	{
		if (A.Name.StartsWith(VH.A(154004)))
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
					return !A.Visible;
				}
			}
		}
		return false;
	}

	internal static Regex A()
	{
		return new Regex(VH.A(41312) + NAME_PATTERN + VH.A(41262), RegexOptions.IgnoreCase);
	}

	internal static void A(Range A)
	{
		try
		{
			A.Worksheet.Activate();
			A.Activate();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static List<Range> B(List<Range> A, Name B)
	{
		List<Range> list = Unapply(B);
		using (List<Range>.Enumerator enumerator = list.GetEnumerator())
		{
			AG aG = default(AG);
			while (enumerator.MoveNext())
			{
				aG = new AG(aG);
				aG.A = enumerator.Current;
				if (A.Find(aG.A) != null)
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
				A.Add(aG.A);
			}
		}
		list = null;
		return A;
	}

	public static void Unapply()
	{
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		checked
		{
			Range range3 = default(Range);
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
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				Collection collection = new Collection();
				Microsoft.Office.Interop.Excel.Application application2 = application;
				if (!(application2.ActiveSheet is Worksheet))
				{
					return;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (!(application2.Selection is Range))
					{
						return;
					}
					Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application2.ActiveWorkbook;
					Worksheet worksheet = (Worksheet)application2.ActiveSheet;
					application2 = null;
					int count = activeWorkbook.Names.Count;
					Range range;
					List<ZF> list;
					ZF item;
					if (count == 0)
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
						Forms.InfoMessage(VH.A(154017));
					}
					else if (count > 1000)
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
						Forms.WarningMessage(VH.A(100155));
					}
					else
					{
						if (count > 100)
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
							if (MessageBox.Show(VH.A(100295), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
							{
								goto IL_0719;
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
						if (worksheet.ProtectContents)
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
							Forms.WarningMessage(VH.A(154094));
						}
						else if (worksheet.TransitionFormEntry)
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
							Forms.WarningMessage(VH.A(154149));
						}
						else
						{
							range = (Range)application.Selection;
							Range range2 = range;
							if (Operators.ConditionalCompareObjectEqual(range2.Cells.CountLarge, 1, TextCompare: false))
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
								if (Conversions.ToBoolean(range2.HasFormula))
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
									range3 = range;
								}
							}
							else
							{
								try
								{
									range3 = range2.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
							}
							range2 = null;
							if (range3 == null)
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
								Forms.InfoMessage(VH.A(154573));
							}
							else
							{
								if (Operators.ConditionalCompareObjectGreater(Operators.MultiplyObject(count, range3.Cells.CountLarge), 100, TextCompare: false))
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
									if (MessageBox.Show(VH.A(154652), VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
									{
										goto IL_0719;
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
								list = new List<ZF>();
								Microsoft.Office.Interop.Excel.Names names = activeWorkbook.Names;
								int num = count;
								for (int i = 1; i <= num; i++)
								{
									Name name = names.Item(i, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									if (!Names.A(name.Name))
									{
										item = new ZF
										{
											A = name.Name,
											B = Conversions.ToString(name.RefersTo)
										};
										object c;
										if (Strings.InStr(item.A, VH.A(7827)) == 0)
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
											c = item.A;
										}
										else
										{
											c = Strings.Split(item.A, VH.A(7827))[1];
										}
										item.C = (string)c;
										item.A = !(name.Parent is Microsoft.Office.Interop.Excel.Workbook);
										item.B = !name.Visible;
										item.D = A(name);
										if (item.D)
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
											item.D = "";
										}
										else
										{
											try
											{
												string d = Conversions.ToString(NewLateBinding.LateGet(name.RefersToRange.Parent, null, VH.A(19019), new object[0], null, null, null));
												item.D = d;
											}
											catch (Exception ex3)
											{
												ProjectData.SetProjectError(ex3);
												Exception ex4 = ex3;
												item.D = "";
												ProjectData.ClearProjectError();
											}
										}
										list.Add(item);
									}
									name = null;
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
								names = null;
								try
								{
									using List<ZF>.Enumerator enumerator = list.GetEnumerator();
									while (enumerator.MoveNext())
									{
										ZF current = enumerator.Current;
										collection.Add(current.C, current.C);
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_0479;
										}
										continue;
										end_IL_0479:
										break;
									}
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									Forms.WarningMessage(VH.A(154934));
									ProjectData.ClearProjectError();
									return;
								}
								collection = null;
								List<ZF> list2 = list;
								Comparison<ZF> comparison;
								if (_Closure_0024__.A == null)
								{
									comparison = (_Closure_0024__.A = [SpecialName] (ZF A, ZF B) => B.A.Length.CompareTo(A.A.Length));
								}
								else
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
									comparison = _Closure_0024__.A;
								}
								list2.Sort(comparison);
								application.ScreenUpdating = false;
								application.EnableEvents = false;
								bool flag = JH.A(range);
								foreach (object item2 in range3)
								{
									object objectValue = RuntimeHelpers.GetObjectValue(item2);
									string b;
									bool f;
									if (Conversions.ToBoolean(NewLateBinding.LateGet(objectValue, null, VH.A(155354), new object[0], null, null, null)))
									{
										b = Conversions.ToString(NewLateBinding.LateGet(objectValue, null, VH.A(58046), new object[0], null, null, null));
										f = true;
									}
									else
									{
										b = Conversions.ToString(NewLateBinding.LateGet(objectValue, null, VH.A(68956), new object[0], null, null, null));
										f = false;
									}
									using List<ZF>.Enumerator enumerator3 = list.GetEnumerator();
									while (enumerator3.MoveNext())
									{
										ZF current2 = enumerator3.Current;
										if (current2.D)
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
										string c2 = Conversions.ToString(NewLateBinding.LateGet(objectValue, null, VH.A(68956), new object[0], null, null, null));
										string a = current2.A;
										string a2 = Strings.Right(current2.B, Strings.Len(current2.B) - 1);
										if (Operators.CompareString(current2.D, worksheet.Name, TextCompare: false) == 0)
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
											a2 = B(current2.B);
										}
										if (current2.A)
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
											a = B(current2.A);
										}
										try
										{
											a2 = A(a2);
										}
										catch (Exception ex7)
										{
											ProjectData.SetProjectError(ex7);
											Exception ex8 = ex7;
											ProjectData.ClearProjectError();
											continue;
										}
										A(a, b, c2, a2, (Range)objectValue, f);
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_06b8;
										}
										continue;
										end_IL_06b8:
										break;
									}
								}
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
									JH.A(range, VH.A(155371));
								}
							}
						}
					}
					goto IL_0719;
					IL_0719:
					application.ScreenUpdating = true;
					application.EnableEvents = true;
					application = null;
					activeWorkbook = null;
					worksheet = null;
					range3 = null;
					range = null;
					list = null;
					item = default(ZF);
					return;
				}
			}
		}
	}

	public static List<Range> Unapply(Name nm)
	{
		List<Range> list = new List<Range>();
		List<Range> list2 = new List<Range>();
		List<Range> result;
		Range refersToRange;
		try
		{
			refersToRange = nm.RefersToRange;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
			goto IL_021e;
		}
		list = GetDependents(refersToRange, nm);
		using (List<Range>.Enumerator enumerator = list.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Range current = enumerator.Current;
				Range range = current;
				string text;
				bool f;
				if (Conversions.ToBoolean(range.HasArray))
				{
					text = Conversions.ToString(range.FormulaArray);
					f = true;
				}
				else
				{
					text = Conversions.ToString(range.Formula);
					f = false;
				}
				range = null;
				Name name = nm;
				string c = Conversions.ToString(current.Formula);
				string a = name.Name;
				string a2 = Strings.Right(Conversions.ToString(name.RefersTo), checked(Strings.Len(RuntimeHelpers.GetObjectValue(name.RefersTo)) - 1));
				try
				{
					if (Operators.CompareString(Conversions.ToString(NewLateBinding.LateGet(name.RefersToRange.Parent, null, VH.A(19019), new object[0], null, null, null)), current.Worksheet.Name, TextCompare: false) == 0)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							a2 = B(Conversions.ToString(name.RefersTo));
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
				name = null;
				if (nm.Parent is Worksheet)
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
					a = B(nm.Name);
				}
				try
				{
					a2 = A(a2);
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
					continue;
				}
				A(a, text, c, a2, current, f);
				try
				{
					if (!Operators.ConditionalCompareObjectNotEqual(current.Formula, text, TextCompare: false))
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
						list2.Add(current);
						break;
					}
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
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_01ff;
				}
				continue;
				end_IL_01ff:
				break;
			}
		}
		refersToRange = null;
		result = list2;
		goto IL_021e;
		IL_021e:
		return result;
	}

	private static string A(string A)
	{
		XlReferenceType xlReferenceType;
		try
		{
			int num;
			if (Conversions.ToInteger(KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(155396)).InnerText) != 0)
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
				num = 1;
			}
			else
			{
				num = 4;
			}
			xlReferenceType = (XlReferenceType)num;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			xlReferenceType = XlReferenceType.xlAbsolute;
			ProjectData.ClearProjectError();
		}
		return Conversions.ToString(MH.A.Application.ConvertFormula(A, XlReferenceStyle.xlA1, XlReferenceStyle.xlA1, xlReferenceType, RuntimeHelpers.GetObjectValue(Missing.Value)));
	}

	private static void A(string A, string B, string C, string D, Range E, bool F)
	{
		try
		{
			if (Strings.InStr(C, A) == 0)
			{
				return;
			}
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
				C = Regex.Replace(C, VH.A(4544) + A + VH.A(4544), D);
				Range range = E;
				object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
				if (F)
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
					range.FormulaArray = C;
				}
				else
				{
					range.Formula = C;
				}
				if (Operators.ConditionalCompareObjectNotEqual(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)), objectValue, TextCompare: false))
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
					if (F)
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
						range.FormulaArray = B;
					}
					else
					{
						range.Formula = B;
					}
				}
				range = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static string B(string A)
	{
		string[] array = Strings.Split(A, VH.A(7827));
		return array[Information.UBound(array)];
	}

	public static List<Range> GetDependents(Range rngRefers, Name nm)
	{
		List<Range> list = new List<Range>();
		Microsoft.Office.Interop.Excel.Application application = rngRefers.Application;
		Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)rngRefers.Worksheet.Parent;
		XlDisplayDrawingObjects displayDrawingObjects = workbook.DisplayDrawingObjects;
		workbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlHide;
		Stopwatch A = new Stopwatch();
		A.Start();
		if (Helpers.ContainsMergedCells(rngRefers))
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
					return list;
				}
			}
		}
		Range range = (Range)rngRefers.Cells[1, 1];
		checked
		{
			try
			{
				try
				{
					range.ShowDependents(RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					if (range.Worksheet.ProtectContents)
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
						MessageBox.Show(ex2.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
					else
					{
						MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
						clsReporting.LogException(ex2);
					}
					throw new Exception();
				}
				int num = 1;
				int num2 = 1;
				bool flag = true;
				string right = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
				while (true)
				{
					application.Goto(range, RuntimeHelpers.GetObjectValue(Missing.Value));
					try
					{
						application.ActiveCell.NavigateArrow(false, num, num2);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
						goto IL_0276;
					}
					if (Operators.CompareString(application.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
					{
						try
						{
							range.NavigateArrow(false, num + 1, num2);
							if (Operators.CompareString(application.ActiveCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_0235;
									}
									continue;
									end_IL_0235:
									break;
								}
								goto IL_0276;
							}
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
							goto IL_0276;
						}
					}
					flag = false;
					list.Add(application.ActiveCell);
					num2++;
					Names.A(ref A, nm);
					continue;
					IL_0276:
					if (flag)
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
					num2 = 1;
					flag = true;
					num++;
					Names.A(ref A, nm);
				}
			}
			catch (TimeoutException ex7)
			{
				ProjectData.SetProjectError(ex7);
				TimeoutException ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
			workbook.DisplayDrawingObjects = displayDrawingObjects;
			try
			{
				NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
			application = null;
			workbook = null;
			range = null;
			A = null;
			return list;
		}
	}

	private static void A(ref Stopwatch A, Name B)
	{
		if (A.Elapsed.Seconds <= KH.A.DependentsTimeout)
		{
			return;
		}
		if (Forms.YesNoMessage(VH.A(155431) + B.Name + VH.A(155556)) == DialogResult.Yes)
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
					throw new TimeoutException();
				}
			}
		}
		A.Restart();
	}
}
