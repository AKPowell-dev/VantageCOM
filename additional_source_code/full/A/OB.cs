using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class OB
{
	internal static void A(Application A, Action<Workbook> B)
	{
		try
		{
			Workbooks workbooks = A.Workbooks;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = workbooks.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Workbook obj = (Workbook)enumerator.Current;
					B(obj);
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
		finally
		{
			Workbook obj = null;
			Workbooks workbooks = null;
		}
	}

	internal static void A(Sheets A, Action<Worksheet> B, Action<long> C = null, Func<Worksheet, bool> D = null)
	{
		List<Worksheet> list = new List<Worksheet>();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					if (!(RuntimeHelpers.GetObjectValue(enumerator.Current) is Worksheet item))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					list.Add(item);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_004d;
					}
					continue;
					end_IL_004d:
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
			if (C != null)
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
				C(list.Count);
			}
			using List<Worksheet>.Enumerator enumerator2 = list.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Worksheet current = enumerator2.Current;
				bool? flag = D?.Invoke(current);
				bool? flag2;
				if (!flag.HasValue)
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
					flag2 = flag;
				}
				else
				{
					flag2 = flag != true;
				}
				flag = flag2;
				if (flag == true)
				{
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
				B(current);
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
			Worksheet current = null;
			list = null;
		}
	}

	internal static void A(Sheets A, Action<Chart> B, Action<long> C = null, Func<Chart, bool> D = null)
	{
		List<Chart> list = new List<Chart>();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					if (RuntimeHelpers.GetObjectValue(enumerator.Current) is Chart item)
					{
						list.Add(item);
					}
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
			C?.Invoke(A.Count);
			using List<Chart>.Enumerator enumerator2 = list.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Chart current = enumerator2.Current;
				bool? obj;
				if (D == null)
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
					obj = null;
				}
				else
				{
					obj = D(current);
				}
				bool? flag = obj;
				if (((!flag) ?? flag) == true)
				{
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
				B(current);
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
			list = null;
		}
	}
}
