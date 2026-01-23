using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class MergedCells : Observation
{
	public MergedCells(Severity sev, Range rng)
		: base(Category.BestPractices, sev, VH.A(20318), rng)
	{
		base.Explanation = VH.A(20343);
		base.HasFix = true;
		base.CanFixMultiple = true;
	}

	public override void FixAction()
	{
		base.FixAction();
		base.Range.UnMerge();
	}
}
