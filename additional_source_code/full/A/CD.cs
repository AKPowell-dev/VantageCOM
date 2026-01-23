using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[StandardModule]
internal sealed class CD
{
	internal static void A(Range A, ref string B, ref string C)
	{
		try
		{
			if (A == null)
			{
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
					break;
				}
			}
			else
			{
				Worksheet worksheet = A.Worksheet;
				if (worksheet.Parent is Workbook workbook)
				{
					B = string.Format(VH.A(48282), workbook.Name);
					C = string.Format(VH.A(48282), worksheet.Name);
					return;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0030;
					}
					continue;
					end_IL_0030:
					break;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Workbook workbook2 = null;
			Worksheet worksheet = null;
		}
		B = "";
		C = "";
	}

	internal static string A(string A, string B, string C)
	{
		if (A == null)
		{
			A = "";
		}
		if (B == null)
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
			B = "";
		}
		if (C == null)
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
			C = "";
		}
		if (Operators.CompareString(A, "", TextCompare: false) != 0)
		{
			B = string.Format(VH.A(49949), A, B);
		}
		if (Operators.CompareString(B, "", TextCompare: false) == 0)
		{
			return C;
		}
		if (Regex.IsMatch(B, VH.A(43285)))
		{
			return string.Format(VH.A(49966), B, C);
		}
		return string.Format(VH.A(49966), B.Replace(VH.A(39851), VH.A(39854)), C);
	}

	internal static string A(Range A, string B, string C)
	{
		string text = A.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		if (!string.IsNullOrEmpty(C))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!string.IsNullOrEmpty(B))
			{
				string B2 = string.Empty;
				string C2 = string.Empty;
				CD.A(A, ref B2, ref C2);
				if (Operators.CompareString(B, B2, TextCompare: false) == 0)
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
					B2 = "";
				}
				if (Operators.CompareString(C, C2, TextCompare: false) == 0)
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
					if (Operators.CompareString(B2, "", TextCompare: false) == 0)
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
						C2 = "";
					}
				}
				return CD.A(B2, C2, text);
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
		return text;
	}
}
