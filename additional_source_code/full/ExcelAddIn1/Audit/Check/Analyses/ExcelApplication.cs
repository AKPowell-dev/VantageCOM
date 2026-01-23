using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class ExcelApplication
{
	[CompilerGenerated]
	internal sealed class N
	{
		public Analysis A;

		public Settings A;

		public Application A;

		[SpecialName]
		internal void A(Application A)
		{
			this.A();
		}

		[SpecialName]
		internal void A()
		{
			Analysis analysis = this.A;
			List<Observation> A = analysis.Observations;
			ExcelApplication.A(ref A, this.A.CalculationModeManual, this.A);
			analysis.Observations = A;
		}

		[SpecialName]
		internal bool A()
		{
			return this.A.CalculationModeManual != Severity.Ignore;
		}
	}

	internal static List<RB> A(Analysis A, Settings B, Application C)
	{
		N CS_0024_003C_003E8__locals5 = new N();
		CS_0024_003C_003E8__locals5.A = A;
		CS_0024_003C_003E8__locals5.A = B;
		CS_0024_003C_003E8__locals5.A = C;
		return new List<RB>
		{
			new TB(VH.A(2479), [SpecialName] (Application application) =>
			{
				CS_0024_003C_003E8__locals5.A();
			}, [SpecialName] () => CS_0024_003C_003E8__locals5.A.CalculationModeManual != Severity.Ignore)
		};
	}

	private static void A(ref List<Observation> A, Severity B, Application C)
	{
		if (B == Severity.Ignore)
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
			if (C.Calculation != XlCalculation.xlCalculationManual)
			{
				return;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				A.Add(new CalcModeManual(B));
				return;
			}
		}
	}
}
