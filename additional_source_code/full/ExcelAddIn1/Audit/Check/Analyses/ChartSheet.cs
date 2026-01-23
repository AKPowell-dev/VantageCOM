using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class ChartSheet
{
	[CompilerGenerated]
	internal sealed class L
	{
		public Analysis A;

		public Settings A;

		[SpecialName]
		internal void A(Chart A)
		{
			Analysis analysis = this.A;
			List<Observation> A2 = analysis.Observations;
			ChartSheet.A(ref A2, this.A.HiddenSheets, A);
			analysis.Observations = A2;
		}

		[SpecialName]
		internal bool A()
		{
			return this.A.HiddenSheets != Severity.Ignore;
		}

		[SpecialName]
		internal void B(Chart A)
		{
			Analysis analysis = this.A;
			List<Observation> A2 = analysis.Observations;
			ChartSheet.B(ref A2, this.A.VeryHiddenSheets, A);
			analysis.Observations = A2;
		}

		[SpecialName]
		internal bool B()
		{
			return this.A.VeryHiddenSheets != Severity.Ignore;
		}
	}

	internal static List<RB> A(Analysis A, Settings B)
	{
		return new List<RB>
		{
			new DC(VH.A(2097), [SpecialName] (Chart c) =>
			{
				Analysis analysis = A;
				List<Observation> A2 = analysis.Observations;
				ChartSheet.A(ref A2, B.HiddenSheets, c);
				analysis.Observations = A2;
			}, [SpecialName] () => B.HiddenSheets != Severity.Ignore),
			new DC(VH.A(2124), [SpecialName] (Chart c) =>
			{
				Analysis analysis = A;
				List<Observation> A2 = analysis.Observations;
				ChartSheet.B(ref A2, B.VeryHiddenSheets, c);
				analysis.Observations = A2;
			}, [SpecialName] () => B.VeryHiddenSheets != Severity.Ignore)
		};
	}

	private static void A(ref List<Observation> A, Severity B, Chart C)
	{
		if (B == Severity.Ignore || C.Visible != XlSheetVisibility.xlSheetHidden)
		{
			return;
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
			A.Add(new HiddenSheet(B, C));
			return;
		}
	}

	private static void B(ref List<Observation> A, Severity B, Chart C)
	{
		if (B != Severity.Ignore && C.Visible == XlSheetVisibility.xlSheetVeryHidden)
		{
			A.Add(new VeryHiddenSheet(B, C));
		}
	}
}
