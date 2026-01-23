using System;
using A;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class FileTooOld : Observation
{
	public FileTooOld(Severity sev, DateTime dtCreated)
		: base(Category.Workbook, sev, VH.A(17117))
	{
		base.Subtitle = VH.A(17134) + dtCreated.ToLocalTime().ToShortDateString();
		base.Explanation = VH.A(17157);
	}
}
