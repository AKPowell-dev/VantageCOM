using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Audit;
using ExcelAddIn1.Audit.Check;
using ExcelAddIn1.Formulas;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class FB
{
	internal sealed class EB
	{
		[CompilerGenerated]
		private List<ParenthesesPair> A;

		[CompilerGenerated]
		private string A;

		internal List<ParenthesesPair> Pairs
		{
			[CompilerGenerated]
			get
			{
				return this.A;
			}
			[CompilerGenerated]
			set
			{
				this.A = value;
			}
		}

		internal string MaskedFormula
		{
			[CompilerGenerated]
			get
			{
				return A;
			}
			[CompilerGenerated]
			set
			{
				A = value;
			}
		}

		internal EB(List<ParenthesesPair> A, string B)
		{
			Pairs = A;
			MaskedFormula = B;
		}
	}

	private readonly Analysis m_A;

	private readonly Dictionary<string, EB> m_A;

	internal FB(Analysis A)
	{
		this.m_A = new Dictionary<string, EB>();
		this.m_A = A;
	}

	internal EB A(Range A)
	{
		EB eB = B(A);
		if (eB == null)
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
			eB = C(A);
		}
		return eB;
	}

	internal int A(Range A)
	{
		EB value = null;
		if (!this.m_A.TryGetValue(FB.A(A), out value))
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
					return 0;
				}
			}
		}
		return value.Pairs.Count;
	}

	private EB B(Range A)
	{
		EB value = null;
		if (!this.m_A.TryGetValue(FB.A(A), out value))
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
					return null;
				}
			}
		}
		return value;
	}

	private EB C(Range A)
	{
		try
		{
			string B = "";
			EB eB = new EB(this.A(A, ref B), B);
			this.A(A, eB);
			return eB;
		}
		finally
		{
		}
	}

	private List<ParenthesesPair> A(Range A, ref string B)
	{
		B = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(QB.A(A), A.Worksheet.Name);
		ExcelAddIn1.Audit.Helpers.MaskQuotedText(ref B);
		if (B.Contains(VH.A(7827)))
		{
			List<Range> list = this.m_A.PrecRetriever.A(A);
			using List<Range>.Enumerator enumerator = list.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range current = enumerator.Current;
				ExcelAddIn1.Audit.Helpers.MaskSheetAndWorkbookNames(ref B, current.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)));
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
		return ExcelAddIn1.Audit.Helpers.IdentifyParenthesesPairs(B);
	}

	private void A(Range A, EB B)
	{
		this.m_A[FB.A(A)] = B;
	}

	internal void A()
	{
		this.m_A.Clear();
	}

	private static string A(Range A)
	{
		return A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
	}
}
