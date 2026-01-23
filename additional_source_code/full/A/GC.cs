using System.Collections.Generic;
using ExcelAddIn1.Audit.Check;
using ExcelAddIn1.Audit.Check.Observations;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class GC
{
	internal static int A(this List<Observation> A)
	{
		return GC.A(A, Severity.High);
	}

	internal static int B(this List<Observation> A)
	{
		return GC.A(A, Severity.Medium);
	}

	internal static int C(this List<Observation> A)
	{
		return GC.A(A, Severity.Low);
	}

	private static int A(List<Observation> A, Severity B)
	{
		int num = 0;
		checked
		{
			using List<Observation>.Enumerator enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Observation current = enumerator.Current;
				if (current.Children.Count > 0)
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
					num += GC.A(current.Children, B);
				}
				else
				{
					if (current.Severity != B)
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
					num++;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				return num;
			}
		}
	}
}
