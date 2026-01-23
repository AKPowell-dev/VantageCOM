using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class NumericIndexReference : Observation
{
	public NumericIndexReference(Severity sev, Range rng, string strFunctionName)
		: base(Category.FormulaIntegrity, sev, VH.A(20963), rng)
	{
		base.Subtitle = strFunctionName + VH.A(9959) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(21010);
	}
}
