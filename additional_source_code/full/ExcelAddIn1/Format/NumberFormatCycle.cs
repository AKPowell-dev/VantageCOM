using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros.Config.Settings;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class NumberFormatCycle
{
	public struct NumberFormat
	{
		public string Name;

		public string Format;
	}

	[CompilerGenerated]
	private List<NumberFormat> A;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private string A;

	public List<NumberFormat> Items
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

	public NumberFormatCycle(XmlDocument xmlSettings, string strNodeName, string strActivity)
	{
		Items = new List<NumberFormat>();
		XmlNodeList xmlNodeList = xmlSettings.DocumentElement.SelectNodes(Constants.XML_NUMBER_FORMATS + VH.A(75498) + strNodeName + VH.A(146515));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = xmlNodeList.GetEnumerator();
			while (enumerator.MoveNext())
			{
				XmlNode xmlNode = (XmlNode)enumerator.Current;
				NumberFormat item = new NumberFormat
				{
					Name = xmlNode.SelectSingleNode(VH.A(19019)).InnerText
				};
				string innerText = xmlNode.SelectSingleNode(VH.A(60221)).InnerText;
				if (Operators.CompareString(innerText, VH.A(20593), TextCompare: false) == 0)
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
					item.Format = "";
				}
				else
				{
					item.Format = innerText;
				}
				Items.Add(item);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		xmlNodeList = null;
		Index = 0;
		Activity = strActivity;
	}
}
