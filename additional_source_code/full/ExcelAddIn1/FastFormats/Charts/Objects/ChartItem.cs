using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public abstract class ChartItem : BaseItem
{
	internal abstract void A(Microsoft.Office.Interop.Excel.Chart A);

	internal new string B(XmlNode A)
	{
		return clsXml.GetAttributeValue(A, FormatConstants.ATTR_TOP);
	}

	internal string C(XmlNode A)
	{
		return clsXml.GetAttributeValue(A, FormatConstants.ATTR_LEFT);
	}
}
