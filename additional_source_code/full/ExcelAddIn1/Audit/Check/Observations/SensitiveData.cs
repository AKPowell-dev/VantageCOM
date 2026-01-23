using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class SensitiveData : Observation
{
	public SensitiveData(Severity sev, Range rng, string strDataType)
		: base(Category.PrivacySecurity, sev, VH.A(22483), rng)
	{
		base.Subtitle = strDataType + VH.A(17350) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(22512);
	}
}
