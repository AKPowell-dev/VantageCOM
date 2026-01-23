using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class CalcModeManual : Observation
{
	public CalcModeManual(Severity sev)
		: base(Category.Data, sev, VH.A(10212))
	{
		base.Subtitle = VH.A(10249);
		base.Explanation = VH.A(10302);
		base.HasFix = true;
	}

	public override void FixAction()
	{
		base.FixAction();
		MH.A.Application.Calculation = XlCalculation.xlCalculationSemiautomatic;
	}
}
