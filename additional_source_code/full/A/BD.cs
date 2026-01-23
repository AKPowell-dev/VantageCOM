using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class BD
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<List<Range>, Range> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal Range A(List<Range> A)
		{
			return BD.A(A);
		}
	}

	internal static bool A(string A, IEnumerable<string> B, Application C, ref List<Range> D)
	{
		List<Range> D2 = null;
		List<List<Range>> list = null;
		try
		{
			if (!BD.B(A, B, C, ref D2))
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
						return false;
					}
				}
			}
			list = BD.A(D2);
			List<List<Range>> source = list;
			Func<List<Range>, Range> selector;
			if (_Closure_0024__.A == null)
			{
				selector = (_Closure_0024__.A = [SpecialName] (List<Range> a) => BD.A(a));
			}
			else
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
				selector = _Closure_0024__.A;
			}
			D = source.Select(selector).ToList();
			return true;
		}
		finally
		{
			if (list != null)
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
				list.Clear();
			}
			list = null;
			if (D2 != null)
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
				D2.Clear();
			}
			D2 = null;
		}
	}

	private static bool B(string A, IEnumerable<string> B, Application C, ref List<Range> D)
	{
		IEnumerator<string> enumerator = default(IEnumerator<string>);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				try
				{
					Workbook workbook = ((!current.Contains(VH.A(43340))) ? C.ActiveWorkbook : C.Workbooks[Strings.Mid(current, 2, checked(Strings.InStr(current, VH.A(43340)) - 2))]);
					Name name = null;
					try
					{
						name = workbook.Names.Item(A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					if (name == null)
					{
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
							break;
						}
						continue;
					}
					if (!name.Visible)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00c3;
							}
							continue;
							end_IL_00c3:
							break;
						}
						continue;
					}
					List<Range> obj = D;
					if (obj == null)
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
					}
					else
					{
						obj.Clear();
					}
					D = null;
					D = BD.A(name);
					using List<Range>.Enumerator enumerator2 = D.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Range current2 = enumerator2.Current;
						if (Operators.CompareString(workbook.Name, C.ActiveWorkbook.Name, TextCompare: false) != 0)
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
							if (Operators.ConditionalCompareObjectEqual(current, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(VH.A(7120), NewLateBinding.LateGet(current2.Worksheet.Parent, null, VH.A(19019), new object[0], null, null, null)), VH.A(43340)), current2.Worksheet.Name), VH.A(7827)), current2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))), TextCompare: false))
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										return true;
									}
								}
							}
						}
						else if (current.Contains(VH.A(7827)))
						{
							if (Operators.CompareString(current, current2.Worksheet.Name + VH.A(7827) + current2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										return true;
									}
								}
							}
						}
						else if (Operators.CompareString(current, current2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									return true;
								}
							}
						}
						current2 = null;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_02d4;
						}
						continue;
						end_IL_02d4:
						break;
					}
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				finally
				{
					Name name = null;
					Workbook workbook = null;
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_030f;
				}
				continue;
				end_IL_030f:
				break;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		return false;
	}

	private static List<Range> A(Name A)
	{
		List<Range> list = new List<Range>();
		List<Range> result;
		try
		{
			Areas areas = A.RefersToRange.Areas;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					list.Add((Range)objectValue);
				}
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
			result = list;
		}
		catch (object obj) when (((Func<bool>)delegate
		{
			// Could not convert BlockContainer to single expression
			COMException obj2 = obj as COMException;
			System.Runtime.CompilerServices.Unsafe.SkipInit(out int result2);
			if (obj2 == null)
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
				result2 = 0;
			}
			else
			{
				ProjectData.SetProjectError(obj2);
				result2 = ((obj2.HResult == -2146827284) ? 1 : 0);
			}
			return (byte)result2 != 0;
		}).Invoke())
		{
			ProjectData.ClearProjectError();
			goto IL_00d4;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = list;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Areas areas = null;
		}
		goto IL_0233;
		IL_00d4:
		checked
		{
			try
			{
				if (!(RuntimeHelpers.GetObjectValue(A.Parent) is Workbook workbook))
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						result = null;
						break;
					}
					goto IL_0233;
				}
				string text = Conversions.ToString(A.RefersTo);
				if (text.StartsWith(VH.A(48936)))
				{
					text = text.Substring(1, text.Length - 1);
				}
				foreach (string item2 in BD.A(text, ','))
				{
					try
					{
						List<string> list2 = BD.A(item2, '!');
						string text2 = list2[0];
						string cell = list2[1];
						if (text2[0] == '\'')
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
							text2 = text2.Substring(1, text2.Length - 2).Replace(VH.A(39854), VH.A(39851));
						}
						Range item = ((_Worksheet)(Worksheet)workbook.Worksheets[text2]).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
						list.Add(item);
					}
					finally
					{
						Range item = null;
					}
				}
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				ProjectData.ClearProjectError();
			}
			finally
			{
				Workbook workbook2 = null;
			}
			result = list;
			goto IL_0233;
		}
		IL_0233:
		return result;
	}

	private static List<List<Range>> A(List<Range> A)
	{
		List<List<Range>> list = new List<List<Range>>();
		List<string> list2 = new List<string>();
		List<string> list3 = new List<string>();
		checked
		{
			try
			{
				using List<Range>.Enumerator enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range current = enumerator.Current;
					try
					{
						Worksheet worksheet = current.Worksheet;
						Workbook workbook = worksheet.Parent as Workbook;
						string name = worksheet.Name;
						string name2 = workbook.Name;
						int num = -1;
						int num2 = list3.Count - 1;
						int num3 = 0;
						while (true)
						{
							if (num3 <= num2)
							{
								if (object.Equals(list3[num3], name) && object.Equals(list2[num3], name2))
								{
									num = num3;
									break;
								}
								num3++;
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
						if (num == -1)
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
							list3.Add(name);
							list2.Add(name2);
							list.Add(new List<Range>());
							num = list3.Count - 1;
						}
						list[num].Add(current);
					}
					finally
					{
						Workbook workbook = null;
						current = null;
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0110;
					}
					continue;
					end_IL_0110:
					break;
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				list.Clear();
				list.Add(A.ToList());
				ProjectData.ClearProjectError();
			}
			return list;
		}
	}

	private static Range A(List<Range> A)
	{
		Range range = null;
		Application application = null;
		try
		{
			using List<Range>.Enumerator enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range current = enumerator.Current;
				try
				{
					if (range == null)
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
							range = current;
							break;
						}
						continue;
					}
					if (application == null)
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
						application = current.Application;
					}
					range = application.Union(range, current, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				finally
				{
					current = null;
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_01bf;
				}
				continue;
				end_IL_01bf:
				break;
			}
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		finally
		{
			application = null;
		}
		return range;
	}

	internal static List<string> A(string A, char B)
	{
		List<string> list = new List<string>();
		bool flag = false;
		StringBuilder stringBuilder = new StringBuilder();
		string text = string.Format(VH.A(49936), A, B);
		foreach (char c in text)
		{
			if (c == B)
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
				if (!flag)
				{
					list.Add(stringBuilder.ToString());
					stringBuilder.Clear();
					continue;
				}
			}
			if (Operators.CompareString(Conversions.ToString(c), VH.A(39851), TextCompare: false) == 0)
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
				flag = !flag;
			}
			stringBuilder.Append(c);
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			return list;
		}
	}
}
