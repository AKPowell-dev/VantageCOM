using System;
using System.Collections;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class DoubleMinus : Observation
{
	public DoubleMinus(Severity sev, Range rng)
		: base(Category.FormulaErrors, sev, VH.A(13781), rng)
	{
		base.Explanation = VH.A(13806);
		base.HasFix = true;
		base.CanFixMultiple = true;
	}

	public override void FixAction()
	{
		base.FixAction();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = base.Range.Cells.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range instance = (Range)enumerator.Current;
				NewLateBinding.LateSet(instance, null, VH.A(1998), new object[1] { NewLateBinding.LateGet(instance, null, VH.A(1998), new object[0], null, null, null).ToString().Replace(VH.A(3799), VH.A(13778)) }, null, null);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}
}
