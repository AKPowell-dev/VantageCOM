using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class CellComment : Observation
{
	public CellComment(Severity sev, Range rng)
		: base(Category.PrivacySecurity, sev, VH.A(10674), rng)
	{
		base.Explanation = VH.A(10699);
		base.HasFix = false;
		base.CanFixMultiple = false;
	}

	public override void FixAction()
	{
		base.FixAction();
		base.Range.ClearComments();
	}
}
