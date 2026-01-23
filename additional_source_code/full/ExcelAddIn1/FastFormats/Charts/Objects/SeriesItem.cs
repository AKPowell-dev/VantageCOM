using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public abstract class SeriesItem : BaseItem
{
	internal abstract void A(Microsoft.Office.Interop.Excel.Series A);
}
