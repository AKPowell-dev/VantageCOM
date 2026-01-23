using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class PB
{
	internal static List<Range> A(this Dictionary<string, List<Range>> A, string B)
	{
		if (!A.ContainsKey(B))
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
			A[B] = new List<Range>();
		}
		return A[B];
	}

	internal static Dictionary<string, List<Range>> A(this Dictionary<string, Dictionary<string, List<Range>>> A, string B)
	{
		if (!A.ContainsKey(B))
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
			A[B] = new Dictionary<string, List<Range>>();
		}
		return A[B];
	}
}
