using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Audit.Check;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class TB : RB
{
	[CompilerGenerated]
	private new Action<Application> m_A;

	internal Action<Application> AppAction
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

	internal TB(string A, Action<Application> B, Func<bool> C = null)
	{
		base.CheckDesc = A;
		AppAction = B;
		base.CondFunc = C;
	}

	internal override void A(Analysis A, Application B = null, Workbook C = null, Sheets D = null)
	{
		AppAction(B);
	}
}
