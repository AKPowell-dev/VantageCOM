using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class LegacyArrayFormula : Observation
{
	public LegacyArrayFormula(Severity sev, Range rng, string strFormula)
		: base(Category.FormulaIntegrity, sev, VH.A(19758), rng)
	{
		base.Subtitle = VH.A(19799) + strFormula + VH.A(19802);
		base.Explanation = VH.A(19805);
	}
}
