using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Tables
{
	public static void IncreaseSize()
	{
		A(A: true);
	}

	public static void DecreaseSize()
	{
		A(A: false);
	}

	private static void A(bool A)
	{
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		checked
		{
			Range range;
			try
			{
				range = JH.A((Range)null);
				bool flag = JH.A(range);
				Range range2 = range;
				int num = Conversions.ToInteger(range2.Rows.CountLarge);
				int num2 = Conversions.ToInteger(range2.Columns.CountLarge);
				float[] array = new float[num - 1 + 1];
				_ = new float[num - 1 + 1];
				int num3 = num - 1;
				for (int i = 0; i <= num3; i++)
				{
					array[i] = Conversions.ToSingle(NewLateBinding.LateGet(range2.Rows[i + 1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(151632), new object[0], null, null, null));
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
					string b;
					if (A)
					{
						Tables.A(range);
						b = VH.A(151651);
					}
					else
					{
						B(range);
						b = VH.A(151690);
					}
					float num4 = 0f;
					int num5 = 0;
					int num6 = num - 1;
					for (int i = 0; i <= num6; i++)
					{
						float num7 = Conversions.ToSingle(Operators.SubtractObject(array[i], NewLateBinding.LateGet(range2.Rows[i + 1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(151632), new object[0], null, null, null)));
						if (num7 > num4)
						{
							num4 = num7;
							num5 = i;
						}
					}
					int num8 = num2 - 1;
					for (int j = 0; j <= num8; j++)
					{
						object instance = range2.Cells[num5 + 1, j + 1];
						NewLateBinding.LateSetComplex(instance, null, VH.A(151729), new object[1] { Operators.DivideObject(Operators.MultiplyObject(NewLateBinding.LateGet(instance, null, VH.A(151632), new object[0], null, null, null), NewLateBinding.LateGet(instance, null, VH.A(151729), new object[0], null, null, null)), array[num5]) }, null, null, OptimisticSet: false, RValueBase: true);
						instance = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						range2 = null;
						if (flag & KH.A.UndoFont)
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
							JH.A(range, b);
						}
						Base.LogActivity(VH.A(151752));
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
			application.ScreenUpdating = true;
			application = null;
			range = null;
		}
	}

	private static void A(Range A)
	{
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Font font;
					(font = ((Range)enumerator.Current).Font).Size = Operators.AddObject(font.Size, 1);
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
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

	private static void B(Range A)
	{
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Font font;
					(font = ((Range)enumerator.Current).Font).Size = Operators.SubtractObject(font.Size, 1);
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
