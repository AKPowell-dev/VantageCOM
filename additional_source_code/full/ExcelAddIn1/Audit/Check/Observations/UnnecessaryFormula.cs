using A;
using ExcelAddIn1.Formulas;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class UnnecessaryFormula : Observation
{
	public UnnecessaryFormula(Severity sev, Range rng)
		: base(Category.FormulaIntegrity, sev, VH.A(9326), rng)
	{
		base.Explanation = VH.A(9365);
		base.HasFix = true;
		base.CanFixMultiple = true;
	}

	public override void FixAction()
	{
		base.FixAction();
		foreach (Range item in base.Range)
		{
			NewLateBinding.LateSet(item, null, VH.A(1998), new object[1] { ExcelAddIn1.Formulas.Helpers.RemoveUnnecessaryEquals(item, Conversions.ToString(NewLateBinding.LateGet(item, null, VH.A(1998), new object[0], null, null, null))) }, null, null);
		}
	}
}
