using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Audit.Check;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class SB : RB
{
	[CompilerGenerated]
	private new Action<Analysis> m_A;

	internal Action<Analysis> FinalAction
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

	internal SB(Action<Analysis> A)
	{
		FinalAction = A;
	}

	internal override void A(Analysis A, Application B = null, Workbook C = null, Sheets D = null)
	{
	}

	internal void B(Analysis A)
	{
		FinalAction(A);
	}
}
