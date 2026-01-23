using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class CellFillColor : Observation
{
	public CellFillColor(Severity sev, Range rng, int intOleColor)
		: base(Category.BrandCompliance, sev, VH.A(10831), rng)
	{
		try
		{
			Color color = ColorTranslator.FromOle(intOleColor);
			base.Subtitle = VH.A(10515) + Conversions.ToString(color.R) + VH.A(10524) + Conversions.ToString(color.G) + VH.A(10524) + Conversions.ToString(color.B) + VH.A(10529) + rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			color = default(Color);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		base.Explanation = VH.A(10862);
	}
}
