using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Audit.Check;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class DC : RB
{
	[CompilerGenerated]
	internal sealed class CC
	{
		public Analysis A;

		[SpecialName]
		internal bool A(Chart A)
		{
			return !this.A.ItemCancelled();
		}
	}

	[CompilerGenerated]
	private new Action<Chart> m_A;

	internal Action<Chart> ChtAction
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

	internal DC(string A, Action<Chart> B, Func<bool> C = null)
	{
		base.CheckDesc = A;
		ChtAction = B;
		base.CondFunc = C;
	}

	internal override void A(Analysis A, Application B = null, Workbook C = null, Sheets D = null)
	{
		OB.A(D, ChtAction, [SpecialName] (long numItems) =>
		{
			base.AssociatedAction.NumItems = numItems;
		}, [SpecialName] (Chart chart) => !A.ItemCancelled());
	}

	[SpecialName]
	[CompilerGenerated]
	private void B(long A)
	{
		base.AssociatedAction.NumItems = A;
	}
}
