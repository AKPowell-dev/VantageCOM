using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class DataOutlier : Observation
{
	public DataOutlier(Severity sev, Range rngOutlier, Range rngData)
		: base(Category.Data, sev, VH.A(12807), rngOutlier)
	{
		base.Subtitle = VH.A(11531) + rngOutlier.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(12832) + rngData.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Explanation = VH.A(12841);
	}
}
