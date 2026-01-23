using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class ShapeOverCells : Observation
{
	public ShapeOverCells(Severity sev, Range rng, Shape shp)
		: base(Category.HiddenData, sev, VH.A(22758), rng)
	{
		base.Shape = shp;
		base.Explanation = VH.A(22791);
	}
}
