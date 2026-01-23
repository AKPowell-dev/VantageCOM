using System;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class UnusedName : Observation
{
	internal UnusedName(Severity A, Name B)
		: base(Category.Performance, A, VH.A(9503))
	{
		try
		{
			base.Range = B.RefersToRange;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		base.Subtitle = VH.A(9526) + B.Name + VH.A(9555);
		base.Explanation = VH.A(9600);
	}
}
