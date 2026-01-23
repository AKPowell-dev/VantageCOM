using System.Xml;
using MacabacusMacros;
using MacabacusMacros.FastFormats.Charts;

namespace ExcelAddIn1.FastFormats.Charts.Objects;

public abstract class BaseItem
{
	public BaseItem()
	{
	}

	internal string B(XmlNode A)
	{
		return clsXml.GetAttributeValue(A, FormatConstants.ATTR_SHOW);
	}
}
