using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class CommentsNotes
{
	internal static void A(WorksheetItem A, object B)
	{
		IEnumerator enumerator = default(IEnumerator);
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
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
					try
					{
						enumerator = ((IEnumerable)NewLateBinding.LateGet((Microsoft.Office.Interop.Excel.Worksheet)B, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
						while (enumerator.MoveNext())
						{
							object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
							A.M((Range)NewLateBinding.LateGet(objectValue, null, VH.A(8701), new object[0], null, null, null));
						}
						return;
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
									break;
								default:
									(enumerator as IDisposable).Dispose();
									goto end_IL_0098;
								}
								continue;
								end_IL_0098:
								break;
							}
						}
					}
				}
			}
		}
		Application application = MH.A.Application;
		foreach (object item in (IEnumerable)NewLateBinding.LateGet(((Range)B).Worksheet, null, VH.A(8668), new object[0], null, null, null))
		{
			object objectValue2 = RuntimeHelpers.GetObjectValue(item);
			if (application.Intersect((Range)B, (Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
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
			A.M((Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null));
		}
		application = null;
	}

	internal static void B(WorksheetItem A, object B)
	{
		CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(B), CommentsNotes.A);
	}

	private static bool A(object A)
	{
		return Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(A, null, VH.A(102617), new object[0], null, null, null)));
	}

	internal static void C(WorksheetItem A, object B)
	{
		CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(B), CommentsNotes.B);
	}

	private static bool B(object A)
	{
		return Conversions.ToBoolean(NewLateBinding.LateGet(A, null, VH.A(102617), new object[0], null, null, null));
	}

	internal static void D(WorksheetItem A, object B)
	{
		CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(B), C);
	}

	private static bool C(object A)
	{
		return CommentsNotes.A(RuntimeHelpers.GetObjectValue(A)) == 0;
	}

	internal static void E(WorksheetItem A, object B)
	{
		CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(B), D);
	}

	private static bool D(object A)
	{
		return CommentsNotes.A(RuntimeHelpers.GetObjectValue(A)) > 0;
	}

	internal static void F(WorksheetItem A, object B)
	{
		CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(B), E);
	}

	private static bool E(object A)
	{
		return CommentsNotes.A(RuntimeHelpers.GetObjectValue(A)) > 4;
	}

	private static void A(WorksheetItem A, object B, Func<object, bool> C)
	{
		IEnumerator enumerator = default(IEnumerator);
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						enumerator = ((IEnumerable)NewLateBinding.LateGet((Microsoft.Office.Interop.Excel.Worksheet)B, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
						while (enumerator.MoveNext())
						{
							object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
							if (C(RuntimeHelpers.GetObjectValue(objectValue)))
							{
								A.M((Range)NewLateBinding.LateGet(objectValue, null, VH.A(8701), new object[0], null, null, null));
							}
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
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									(enumerator as IDisposable).Dispose();
									goto end_IL_00b2;
								}
								continue;
								end_IL_00b2:
								break;
							}
						}
					}
				}
			}
		}
		Application application = MH.A.Application;
		IEnumerator enumerator2 = ((IEnumerable)NewLateBinding.LateGet(((Range)B).Worksheet, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
		try
		{
			while (enumerator2.MoveNext())
			{
				object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
				if (application.Intersect((Range)B, (Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null || !C(RuntimeHelpers.GetObjectValue(objectValue2)))
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
				A.M((Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null));
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_02ea;
				}
				continue;
				end_IL_02ea:
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
		application = null;
	}

	internal static void G(WorksheetItem A, object B)
	{
		Regex c = new Regex(VH.A(4544) + Regex.Escape(Props.SearchForm.Input1.Trim()) + VH.A(4544), RegexOptions.IgnoreCase);
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			foreach (object item in (IEnumerable)NewLateBinding.LateGet((Microsoft.Office.Interop.Excel.Worksheet)B, null, VH.A(8668), new object[0], null, null, null))
			{
				object objectValue = RuntimeHelpers.GetObjectValue(item);
				CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(objectValue), c);
			}
		}
		else
		{
			Application application = MH.A.Application;
			{
				IEnumerator enumerator2 = ((IEnumerable)NewLateBinding.LateGet(((Range)B).Worksheet, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
				try
				{
					while (enumerator2.MoveNext())
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
						if (application.Intersect((Range)B, (Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
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
						CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(objectValue2), c);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_02bb;
						}
						continue;
						end_IL_02bb:
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
			application = null;
		}
		c = null;
	}

	private static void A(WorksheetItem A, object B, Regex C)
	{
		if (C.IsMatch(Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(B, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null))))
		{
			A.M((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null));
			return;
		}
		int num = CommentsNotes.A(RuntimeHelpers.GetObjectValue(B));
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			object[] array;
			bool[] array2;
			object instance = NewLateBinding.LateGet(B, null, VH.A(102647), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
			if (array2[0])
			{
				i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
			}
			if (!C.IsMatch(Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null))))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				A.M((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null));
				return;
			}
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

	internal static void H(WorksheetItem A, object B)
	{
		string c = Props.SearchForm.Input1.ToLower();
		IEnumerator enumerator = default(IEnumerator);
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
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
					try
					{
						enumerator = ((IEnumerable)NewLateBinding.LateGet((Microsoft.Office.Interop.Excel.Worksheet)B, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
						while (enumerator.MoveNext())
						{
							object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
							CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(objectValue), c);
						}
						return;
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
									break;
								default:
									(enumerator as IDisposable).Dispose();
									goto end_IL_0098;
								}
								continue;
								end_IL_0098:
								break;
							}
						}
					}
				}
			}
		}
		Application application = MH.A.Application;
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = ((IEnumerable)NewLateBinding.LateGet(((Range)B).Worksheet, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
			while (enumerator2.MoveNext())
			{
				object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
				if (application.Intersect((Range)B, (Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
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
					break;
				}
				CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(objectValue2), c);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_029c;
				}
				continue;
				end_IL_029c:
				break;
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		application = null;
	}

	private static void A(WorksheetItem A, object B, string C)
	{
		if (CommentsNotes.A(RuntimeHelpers.GetObjectValue(B), C))
		{
			A.M((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null));
			return;
		}
		int num = CommentsNotes.A(RuntimeHelpers.GetObjectValue(B));
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			object[] array;
			bool[] array2;
			object obj = NewLateBinding.LateGet(B, null, VH.A(102647), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
			if (array2[0])
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
				i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
			}
			if (CommentsNotes.A(RuntimeHelpers.GetObjectValue(obj), C))
			{
				A.M((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null));
				return;
			}
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

	private static bool A(object A, string B)
	{
		return NewLateBinding.LateGet(A, null, VH.A(96399), new object[0], null, null, null).ToString().ToLower()
			.Contains(B);
	}

	internal static void I(WorksheetItem A, object B)
	{
		if (!int.TryParse(Props.SearchForm.Input1, out var result))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			try
			{
				enumerator = ((IEnumerable)NewLateBinding.LateGet((Microsoft.Office.Interop.Excel.Worksheet)B, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(objectValue), result);
				}
				return;
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
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							(enumerator as IDisposable).Dispose();
							goto end_IL_008b;
						}
						continue;
						end_IL_008b:
						break;
					}
				}
			}
		}
		Application application = MH.A.Application;
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = ((IEnumerable)NewLateBinding.LateGet(((Range)B).Worksheet, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
			while (enumerator2.MoveNext())
			{
				object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
				if (application.Intersect((Range)B, (Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
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
				CommentsNotes.A(A, RuntimeHelpers.GetObjectValue(objectValue2), result);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_02a2;
				}
				continue;
				end_IL_02a2:
				break;
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		application = null;
	}

	private static void A(WorksheetItem A, object B, int C)
	{
		if (CommentsNotes.A(RuntimeHelpers.GetObjectValue(B), C))
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
					A.M((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null));
					return;
				}
			}
		}
		int num = CommentsNotes.A(RuntimeHelpers.GetObjectValue(B));
		for (int i = 1; i <= num; i = checked(i + 1))
		{
			object[] array;
			bool[] array2;
			object obj = NewLateBinding.LateGet(B, null, VH.A(102647), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
			if (array2[0])
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
				i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
			}
			if (CommentsNotes.A(RuntimeHelpers.GetObjectValue(obj), C))
			{
				A.M((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null));
				break;
			}
		}
	}

	private static bool A(object A, int B)
	{
		return DateTime.Compare(((DateTime)NewLateBinding.LateGet(A, null, VH.A(102662), new object[0], null, null, null)).ToUniversalTime(), DateTime.UtcNow.AddDays(checked(-1 * B))) >= 0;
	}

	internal static void J(WorksheetItem A, object B)
	{
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			IEnumerator enumerator = ((IEnumerable)NewLateBinding.LateGet((Microsoft.Office.Interop.Excel.Worksheet)B, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					K(A, RuntimeHelpers.GetObjectValue(objectValue));
				}
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
		Application application = MH.A.Application;
		IEnumerator enumerator2 = ((IEnumerable)NewLateBinding.LateGet(((Range)B).Worksheet, null, VH.A(8668), new object[0], null, null, null)).GetEnumerator();
		try
		{
			while (enumerator2.MoveNext())
			{
				object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator2.Current);
				if (application.Intersect((Range)B, (Range)NewLateBinding.LateGet(objectValue2, null, VH.A(8701), new object[0], null, null, null), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
				{
					K(A, RuntimeHelpers.GetObjectValue(objectValue2));
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_026d;
				}
				continue;
				end_IL_026d:
				break;
			}
		}
		finally
		{
			IDisposable disposable2 = enumerator2 as IDisposable;
			if (disposable2 != null)
			{
				disposable2.Dispose();
			}
		}
		application = null;
	}

	private static void K(WorksheetItem A, object B)
	{
		if (!RangeHelpers.B((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null)))
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
			A.M((Range)NewLateBinding.LateGet(B, null, VH.A(8701), new object[0], null, null, null));
			return;
		}
	}

	private static int A(object A)
	{
		return Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
	}

	internal static void L(WorksheetItem A, object B)
	{
		IEnumerator enumerator = default(IEnumerator);
		if (B is Microsoft.Office.Interop.Excel.Worksheet)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					{
						enumerator = ((Microsoft.Office.Interop.Excel.Worksheet)B).Comments.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Comment comment = (Comment)enumerator.Current;
								A.L((Range)comment.Parent);
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
		}
		Range range = RangeHelpers.F((Range)B);
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator2 = default(IEnumerator);
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			enumerator2 = range.GetEnumerator();
			try
			{
				while (enumerator2.MoveNext())
				{
					Range a = (Range)enumerator2.Current;
					A.L(a);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_00c8;
					}
					continue;
					end_IL_00c8:
					break;
				}
			}
			finally
			{
				IDisposable disposable2 = enumerator2 as IDisposable;
				if (disposable2 != null)
				{
					disposable2.Dispose();
				}
			}
			range = null;
			return;
		}
	}
}
