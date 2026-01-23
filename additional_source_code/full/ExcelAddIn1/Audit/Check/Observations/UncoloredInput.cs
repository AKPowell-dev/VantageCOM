using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class UncoloredInput : Observation
{
	public UncoloredInput(Severity sev, Range rng)
		: base(Category.BestPractices, sev, VH.A(24841), rng)
	{
		base.Explanation = VH.A(24872);
		base.HasFix = true;
		base.CanFixMultiple = true;
	}

	public override void FixAction()
	{
		base.FixAction();
		base.Range.Font.Color = clsColors.RGB2Ole(KH.A.AutoColors[0]);
	}
}
