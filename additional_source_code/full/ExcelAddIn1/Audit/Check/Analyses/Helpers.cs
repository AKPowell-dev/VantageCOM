using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class Helpers
{
	internal static MatchCollection A(string A, string B)
	{
		return Regex.Matches(A, VH.A(4544) + B + VH.A(4549), RegexOptions.IgnoreCase);
	}

	internal static Range A(Microsoft.Office.Interop.Excel.Worksheet A, string B)
	{
		object objectValue = RuntimeHelpers.GetObjectValue(A.Evaluate(B));
		if (objectValue is Range)
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
					return (Range)objectValue;
				}
			}
		}
		return null;
	}

	internal static string[] A(Match A)
	{
		return A.Groups[1].Value.Split(',');
	}
}
