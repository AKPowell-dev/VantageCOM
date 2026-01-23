using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Config.Settings;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class ColorCycle
{
	public struct Color
	{
		public string RGB;

		public int OLE;

		public XlPattern Pattern;

		public int PatternOLE;
	}

	[CompilerGenerated]
	private List<Color> A;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private string A;

	public List<Color> Colors
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public int Index
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public string Activity
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public ColorCycle(XmlDocument xmlSettings, string strNodeName, string strActivity)
	{
		Colors = new List<Color>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = xmlSettings.DocumentElement.SelectNodes(Constants.XML_COLOR_CYCLES + VH.A(75498) + strNodeName + VH.A(146452)).GetEnumerator();
			while (enumerator.MoveNext())
			{
				XmlNode xmlNode = (XmlNode)enumerator.Current;
				Color item = new Color
				{
					RGB = xmlNode.InnerText
				};
				if (item.RGB.Length > 0)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					item.OLE = clsColors.RGB2Ole(item.RGB);
				}
				else
				{
					item.OLE = -4142;
				}
				if (xmlNode.Attributes[VH.A(146465)] != null)
				{
					item.Pattern = (XlPattern)Conversions.ToInteger(xmlNode.Attributes[VH.A(146465)].Value);
				}
				else
				{
					item.Pattern = XlPattern.xlPatternNone;
				}
				if (xmlNode.Attributes[VH.A(146490)] != null)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					item.PatternOLE = clsColors.RGB2Ole(xmlNode.Attributes[VH.A(146490)].Value);
				}
				else
				{
					item.PatternOLE = 0;
				}
				Colors.Add(item);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0185;
				}
				continue;
				end_IL_0185:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		Index = 0;
		Activity = strActivity;
	}
}
