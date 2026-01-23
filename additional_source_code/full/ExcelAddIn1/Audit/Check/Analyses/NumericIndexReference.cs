using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class NumericIndexReference
{
	internal static void A(ref List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
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
			string c = Conversions.ToString(NewLateBinding.LateGet(C, null, VH.A(1998), new object[0], null, null, null));
			NumericIndexReference.A(ref A, C, c, VH.A(2015));
			NumericIndexReference.A(ref A, C, c, VH.A(2030));
			NumericIndexReference.A(ref A, C, c);
			NumericIndexReference.B(ref A, C, c);
			return;
		}
	}

	private static void A(ref List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> A, Range B, string C, string D)
	{
		bool flag = false;
		if (!C.Contains(D))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			enumerator = Helpers.A(C, D).GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					if (int.TryParse(Helpers.A((Match)enumerator.Current)[2], out var _))
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								flag = true;
								return;
							}
						}
					}
					if (!flag)
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
						A.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference(B, D));
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

	private static void A(ref List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> A, Range B, string C)
	{
		string text = VH.A(4444);
		bool flag = false;
		if (!C.Contains(text))
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
			foreach (Match item in Helpers.A(C, text))
			{
				string[] array = Helpers.A(item);
				int num = array.Length;
				int num2 = 1;
				while (true)
				{
					if (num2 <= num)
					{
						if (int.TryParse(array[num2], out var _))
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
							flag = true;
							break;
						}
						num2 = checked(num2 + 1);
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
					break;
				}
				if (!flag)
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
					A.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference(B, text));
					return;
				}
			}
			return;
		}
	}

	private static void B(ref List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> A, Range B, string C)
	{
		string text = VH.A(4576);
		bool flag = false;
		if (!C.Contains(text))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = Helpers.A(C, text).GetEnumerator();
			while (enumerator.MoveNext())
			{
				string[] array = Helpers.A((Match)enumerator.Current);
				int num = array.Length;
				int num2 = 1;
				while (true)
				{
					if (num2 <= num)
					{
						if (int.TryParse(array[num2], out var _))
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
							flag = true;
							break;
						}
						num2 = checked(num2 + 1);
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
					break;
				}
				if (!flag)
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
					A.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference(B, text));
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
}
