using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Values
{
	[CompilerGenerated]
	internal sealed class EF
	{
		public Func<int, bool> A;

		public Dictionary<Range, object> A;

		public EF(EF A)
		{
			if (A != null)
			{
				this.A = A.A;
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(KeyValuePair<Range, object> A)
		{
			FF fF = new FF(fF)
			{
				A = A
			};
			return this.A(this.A.Where(fF.A).Count());
		}
	}

	[CompilerGenerated]
	internal sealed class FF
	{
		public KeyValuePair<Range, object> A;

		public FF(FF A)
		{
			if (A == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(KeyValuePair<Range, object> A)
		{
			int num;
			if (A.Value.GetType().Equals(this.A.Value.GetType()))
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
				num = (Conversions.ToBoolean(Operators.CompareObjectEqual(A.Value, this.A.Value, TextCompare: false)) ? 1 : 0);
			}
			else
			{
				num = 0;
			}
			return Conversions.ToBoolean((byte)num != 0);
		}
	}

	internal static void A(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<object, object, bool>)Values.A);
	}

	private static bool A(object A, object B)
	{
		return Operators.ConditionalCompareObjectEqual(A, B, TextCompare: false);
	}

	internal static void B(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<object, object, bool>)Values.B);
	}

	private static bool B(object A, object B)
	{
		return Operators.ConditionalCompareObjectNotEqual(A, B, TextCompare: false);
	}

	private static void A(WorksheetItem A, Range B, Func<object, object, bool> C)
	{
		B = RangeHelpers.H(B);
		if (B == null)
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
			Range A2 = null;
			string input = Props.SearchForm.Input1;
			object obj;
			if (Versioned.IsNumeric(input))
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
				obj = Conversions.ToDouble(input);
			}
			else
			{
				obj = input;
			}
			Values.A(ref A2, B, RuntimeHelpers.GetObjectValue(obj), C);
			Q(A, A2);
			A2 = null;
			return;
		}
	}

	internal static void C(WorksheetItem A, Range B)
	{
		Values.B(A, B, C);
	}

	private static bool C(object A, object B)
	{
		return Operators.ConditionalCompareObjectGreater(A, B, TextCompare: false);
	}

	internal static void D(WorksheetItem A, Range B)
	{
		Values.B(A, B, D);
	}

	private static bool D(object A, object B)
	{
		return Operators.ConditionalCompareObjectLess(A, B, TextCompare: false);
	}

	internal static void E(WorksheetItem A, Range B)
	{
		Values.B(A, B, E);
	}

	private static bool E(object A, object B)
	{
		return Operators.ConditionalCompareObjectGreaterEqual(A, B, TextCompare: false);
	}

	internal static void F(WorksheetItem A, Range B)
	{
		Values.B(A, B, F);
	}

	private static bool F(object A, object B)
	{
		return Operators.ConditionalCompareObjectLessEqual(A, B, TextCompare: false);
	}

	private static void B(WorksheetItem A, Range B, Func<object, object, bool> C)
	{
		B = RangeHelpers.CellsWithNumbers(B);
		if (B == null)
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
			Range A2 = null;
			string input = Props.SearchForm.Input1;
			if (!Versioned.IsNumeric(input))
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
				Values.A(ref A2, B, Conversions.ToDouble(input), C);
				Q(A, A2);
				A2 = null;
				return;
			}
		}
	}

	private static void A(ref Range A, Range B, object C, Func<object, object, bool> D)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = B.Areas.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
					object objectValue2 = RuntimeHelpers.GetObjectValue(range.Value2);
					if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
						try
						{
							if (Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)))
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
								if (!D(RuntimeHelpers.GetObjectValue(objectValue2), RuntimeHelpers.GetObjectValue(C)))
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
									RangeHelpers.A(ref A, range);
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
						continue;
					}
					int num = Information.LBound((Array)objectValue);
					int num2 = Information.UBound((Array)objectValue);
					for (int i = num; i <= num2; i++)
					{
						int num3 = Information.LBound((Array)objectValue, 2);
						int num4 = Information.UBound((Array)objectValue, 2);
						for (int j = num3; j <= num4; j++)
						{
							try
							{
								if (Information.IsDate(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null))))
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
									if (D(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue2, new object[2] { i, j }, null)), RuntimeHelpers.GetObjectValue(C)))
									{
										RangeHelpers.A(ref A, (Range)range.Cells[i, j]);
									}
									break;
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
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_01e0;
							}
							continue;
							end_IL_01e0:
							break;
						}
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
		}
	}

	internal static void G(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<object, object, object, bool>)Values.A);
	}

	private static bool A(object A, object B, object C)
	{
		int num;
		if (Conversions.ToBoolean(Operators.CompareObjectGreater(A, B, TextCompare: false)))
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
			num = (Conversions.ToBoolean(Operators.CompareObjectLess(A, C, TextCompare: false)) ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		return Conversions.ToBoolean((byte)num != 0);
	}

	internal static void H(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<object, object, object, bool>)Values.B);
	}

	private static bool B(object A, object B, object C)
	{
		int num;
		if (!Conversions.ToBoolean(Operators.CompareObjectLessEqual(A, B, TextCompare: false)))
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
			num = (Conversions.ToBoolean(Operators.CompareObjectGreaterEqual(A, C, TextCompare: false)) ? 1 : 0);
		}
		else
		{
			num = 1;
		}
		return Conversions.ToBoolean((byte)num != 0);
	}

	private static void A(WorksheetItem A, Range B, Func<object, object, object, bool> C)
	{
		B = RangeHelpers.CellsWithNumbers(B);
		if (B == null)
		{
			return;
		}
		checked
		{
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
				Range A2 = null;
				string input = Props.SearchForm.Input1;
				string input2 = Props.SearchForm.Input2;
				if (!Versioned.IsNumeric(input))
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
					if (!Versioned.IsNumeric(input2))
					{
						return;
					}
					object obj = Conversions.ToDouble(input);
					object obj2 = Conversions.ToDouble(input2);
					try
					{
						enumerator = B.Areas.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Range range = (Range)enumerator.Current;
							object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
							object objectValue2 = RuntimeHelpers.GetObjectValue(range.Value2);
							if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
								try
								{
									if (Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)) || !C(RuntimeHelpers.GetObjectValue(objectValue2), RuntimeHelpers.GetObjectValue(obj), RuntimeHelpers.GetObjectValue(obj2)))
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
										RangeHelpers.A(ref A2, range);
										break;
									}
									continue;
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
								continue;
							}
							int num = Information.LBound((Array)objectValue);
							int num2 = Information.UBound((Array)objectValue);
							for (int i = num; i <= num2; i++)
							{
								int num3 = Information.LBound((Array)objectValue, 2);
								int num4 = Information.UBound((Array)objectValue, 2);
								for (int j = num3; j <= num4; j++)
								{
									try
									{
										if (Information.IsDate(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null))))
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
											if (!C(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue2, new object[2] { i, j }, null)), RuntimeHelpers.GetObjectValue(obj), RuntimeHelpers.GetObjectValue(obj2)))
											{
												break;
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												RangeHelpers.A(ref A2, (Range)range.Cells[i, j]);
												break;
											}
											break;
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
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_0278;
									}
									continue;
									end_IL_0278:
									break;
								}
							}
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_029f;
							}
							continue;
							end_IL_029f:
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
					Q(A, A2);
					A2 = null;
					return;
				}
			}
		}
	}

	internal static void I(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<object, bool>)C);
	}

	private static bool C(object A)
	{
		return Operators.ConditionalCompareObjectEqual(Operators.ModObject(A, 2), 1, TextCompare: false);
	}

	internal static void J(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<object, bool>)D);
	}

	private static bool D(object A)
	{
		return Operators.ConditionalCompareObjectEqual(Operators.ModObject(A, 2), 0, TextCompare: false);
	}

	private static void A(WorksheetItem A, Range B, Func<object, bool> C)
	{
		Range A2 = null;
		B = RangeHelpers.CellsWithNumbers(B);
		if (B == null)
		{
			return;
		}
		checked
		{
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
				try
				{
					enumerator = B.Areas.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						object objectValue = RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
						object objectValue2 = RuntimeHelpers.GetObjectValue(range.Value2);
						if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
							try
							{
								if (Information.IsDate(RuntimeHelpers.GetObjectValue(objectValue)))
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
									if (!C(RuntimeHelpers.GetObjectValue(objectValue2)))
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
										RangeHelpers.A(ref A2, range);
										break;
									}
									break;
								}
								continue;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							continue;
						}
						int num = Information.LBound((Array)objectValue);
						int num2 = Information.UBound((Array)objectValue);
						for (int i = num; i <= num2; i++)
						{
							int num3 = Information.LBound((Array)objectValue, 2);
							int num4 = Information.UBound((Array)objectValue, 2);
							for (int j = num3; j <= num4; j++)
							{
								try
								{
									if (Information.IsDate(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null))) || !C(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue2, new object[2] { i, j }, null))))
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
										RangeHelpers.A(ref A2, (Range)range.Cells[i, j]);
										break;
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
									goto end_IL_01e8;
								}
								continue;
								end_IL_01e8:
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
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				Q(A, A2);
				A2 = null;
				return;
			}
		}
	}

	internal static void K(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<int, bool>)Values.A);
	}

	private static bool A(int A)
	{
		return A == 1;
	}

	internal static void L(WorksheetItem A, Range B)
	{
		Values.A(A, B, (Func<int, bool>)Values.B);
	}

	private static bool B(int A)
	{
		return A > 1;
	}

	private static void A(WorksheetItem A, Range B, Func<int, bool> C)
	{
		EF a = default(EF);
		EF CS_0024_003C_003E8__locals7 = new EF(a);
		CS_0024_003C_003E8__locals7.A = C;
		Range A2 = null;
		CS_0024_003C_003E8__locals7.A = new Dictionary<Range, object>();
		B = RangeHelpers.H(B);
		if (B != null)
		{
			{
				IEnumerator enumerator = B.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						try
						{
							CS_0024_003C_003E8__locals7.A.Add(range, RuntimeHelpers.GetObjectValue(range.Value2));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
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
					IDisposable disposable = enumerator as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
			}
			IEnumerable<KeyValuePair<Range, object>> enumerable = CS_0024_003C_003E8__locals7.A.Where([SpecialName] (KeyValuePair<Range, object> a2) =>
			{
				FF fF = new FF(fF);
				fF.A = a2;
				return CS_0024_003C_003E8__locals7.A(CS_0024_003C_003E8__locals7.A.Where(fF.A).Count());
			});
			IEnumerator<KeyValuePair<Range, object>> enumerator2 = default(IEnumerator<KeyValuePair<Range, object>>);
			try
			{
				enumerator2 = enumerable.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					RangeHelpers.A(ref A2, enumerator2.Current.Key);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_00ee;
					}
					continue;
					end_IL_00ee:
					break;
				}
			}
			finally
			{
				if (enumerator2 != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						enumerator2.Dispose();
						break;
					}
				}
			}
			enumerable = null;
			Q(A, A2);
			A2 = null;
		}
		CS_0024_003C_003E8__locals7.A = null;
	}

	internal static void M(WorksheetItem A, Range B)
	{
		Range range = null;
		try
		{
			range = RangeHelpers.B(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			int num = 0;
			try
			{
				enumerator = range.GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Range a = (Range)enumerator.Current;
						num = checked(num + 1);
						if (num > 25)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A.A(Operators.AddObject(Operators.SubtractObject(range.CountLarge, num), 1).ToString() + VH.A(103928));
								break;
							}
							break;
						}
						A.J(a);
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_00b3;
						}
						continue;
						end_IL_00b3:
						break;
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
			range = null;
			return;
		}
	}

	internal static void N(WorksheetItem A, Range B)
	{
		Values.A(A, B, G);
	}

	private static bool G(double A, double B)
	{
		return A > B;
	}

	internal static void O(WorksheetItem A, Range B)
	{
		Values.A(A, B, H);
	}

	private static bool H(double A, double B)
	{
		return A < B;
	}

	private static void A(WorksheetItem A, Range B, Func<double, double, bool> C)
	{
		B = RangeHelpers.CellsWithNumbers(B);
		if (B == null)
		{
			return;
		}
		checked
		{
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
				Range A2 = null;
				try
				{
					enumerator = B.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range = (Range)enumerator.Current;
						if (A2 != null)
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
							if (MH.A.Application.Intersect(A2, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
						}
						if (Operators.ConditionalCompareObjectEqual(range.CurrentRegion.Cells.CountLarge, 1, TextCompare: false))
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
						long num = 0L;
						long num2 = 0L;
						double? num3 = null;
						object objectValue = RuntimeHelpers.GetObjectValue(range.CurrentRegion.Value2);
						int num4 = Information.LBound((Array)objectValue);
						int num5 = Information.UBound((Array)objectValue);
						for (int i = num4; i <= num5; i++)
						{
							int num6 = Information.LBound((Array)objectValue, 2);
							int num7 = Information.UBound((Array)objectValue, 2);
							for (int j = num6; j <= num7; j++)
							{
								try
								{
									object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(objectValue, new object[2] { i, j }, null));
									if (!(objectValue2 is double))
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_028b;
											}
											continue;
											end_IL_028b:
											break;
										}
										continue;
									}
									if (!num3.HasValue)
									{
										goto IL_02c7;
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
									if (C(Conversions.ToDouble(objectValue2), num3.Value))
									{
										goto IL_02c7;
									}
									goto end_IL_0258;
									IL_02c7:
									num3 = (double?)objectValue2;
									num = i;
									num2 = j;
									end_IL_0258:;
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_02f7;
								}
								continue;
								end_IL_02f7:
								break;
							}
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
						if (num > 0)
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
							A.B((Range)range.CurrentRegion.Cells[num, num2]);
						}
						RangeHelpers.A(ref A2, range.CurrentRegion);
					}
					return;
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
			}
		}
	}

	internal static void P(WorksheetItem A, Range B)
	{
		throw new NotImplementedException();
	}

	private static void Q(WorksheetItem A, Range B)
	{
		if (B == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.Rows.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range a = (Range)enumerator.Current;
				A.B(a);
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
	}

	private static void R(WorksheetItem A, Range B)
	{
		if (B == null)
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
			enumerator = B.Areas.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Range a = (Range)enumerator.Current;
					A.B(a);
				}
				while (true)
				{
					switch (3)
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
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
	}
}
