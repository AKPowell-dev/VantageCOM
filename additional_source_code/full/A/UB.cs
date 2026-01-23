using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Audit.Check;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class UB : RB
{
	[CompilerGenerated]
	private new Action<Workbook> m_A;

	internal Action<Workbook> WbAction
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal UB(string A, Action<Workbook> B, Func<bool> C = null)
	{
		base.CheckDesc = A;
		WbAction = B;
		base.CondFunc = C;
	}

	internal override void A(Analysis A, Application B = null, Workbook C = null, Sheets D = null)
	{
		WbAction(C);
	}
}
