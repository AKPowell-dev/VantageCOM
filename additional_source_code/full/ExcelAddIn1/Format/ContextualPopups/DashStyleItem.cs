using System.Xml;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format.ContextualPopups;

public sealed class DashStyleItem
{
	private MsoLineDashStyle m_A;

	public MsoLineDashStyle OfficeDashStyle
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public XlLineStyle ExcelLineStyle => A();

	public DashStyleItem(XmlNode nd, bool blnChecked)
	{
		OfficeDashStyle = (MsoLineDashStyle)Conversions.ToInteger(nd.Attributes[VH.A(73267)].Value);
	}

	private XlLineStyle A()
	{
		MsoLineDashStyle officeDashStyle = OfficeDashStyle;
		XlLineStyle result = default(XlLineStyle);
		if (officeDashStyle != MsoLineDashStyle.msoLineSolid)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (officeDashStyle != MsoLineDashStyle.msoLineDash)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								return result;
							}
						}
					}
					return XlLineStyle.xlDash;
				}
			}
		}
		return XlLineStyle.xlContinuous;
	}
}
