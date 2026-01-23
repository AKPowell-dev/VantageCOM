using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class CellNote : Observation
{
	public CellNote(Severity sev, Range rng)
		: base(Category.PrivacySecurity, sev, VH.A(10994), rng)
	{
		base.Explanation = VH.A(11013);
		base.HasFix = false;
		base.CanFixMultiple = false;
	}

	public override void FixAction()
	{
		base.FixAction();
		base.Range.ClearComments();
	}
}
