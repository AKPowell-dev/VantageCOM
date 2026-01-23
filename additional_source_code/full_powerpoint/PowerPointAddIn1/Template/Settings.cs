using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Template;

public sealed class Settings
{
	public struct Margins
	{
		public float Top;

		public float Bottom;

		public float Left;

		public float Right;
	}

	[CompilerGenerated]
	private List<string> m_A;

	[CompilerGenerated]
	private List<int> m_A;

	[CompilerGenerated]
	private List<string> m_B;

	[CompilerGenerated]
	private int? m_A;

	[CompilerGenerated]
	private int? m_B;

	[CompilerGenerated]
	private Margins? m_A;

	[CompilerGenerated]
	private Margins? m_B;

	internal List<string> LegalFontTypes
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal List<int> LegalFontSizes
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal List<string> LegalFontColors
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal int? MinFontSize
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal int? MaxFontSize
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal Margins? TextboxMargins
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal Margins? SlideMargins
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal Settings(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		LegalFontTypes = BrandCompliance.GetLegalFontTypes(A);
		LegalFontSizes = null;
		LegalFontColors = null;
		MinFontSize = BrandCompliance.GetMinFontSize(A);
		MaxFontSize = BrandCompliance.GetMaxFontSize(A);
		TextboxMargins = this.A(A);
		SlideMargins = B(A);
	}

	private Margins? A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		XmlDocument xmlDocument = new XmlDocument();
		Margins? result;
		try
		{
			xmlDocument.LoadXml(BrandCompliance.GetXmlFromTags(A));
			XmlNode xmlNode = xmlDocument.DocumentElement.SelectSingleNode(BrandCompliance.XML_TEXTBOX_MGNS);
			result = new Margins
			{
				Top = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120462)).InnerText),
				Bottom = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120469)).InnerText),
				Right = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120482)).InnerText),
				Left = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120493)).InnerText)
			};
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		finally
		{
			XmlNode xmlNode = null;
			xmlDocument = null;
		}
		return result;
	}

	private Margins? B(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		XmlDocument xmlDocument = new XmlDocument();
		Margins? result;
		try
		{
			xmlDocument.LoadXml(BrandCompliance.GetXmlFromTags(A));
			XmlNode xmlNode = xmlDocument.DocumentElement.SelectSingleNode(BrandCompliance.XML_SLIDE_MGNS);
			result = new Margins
			{
				Top = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120462)).InnerText),
				Bottom = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120469)).InnerText),
				Right = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120482)).InnerText),
				Left = Conversions.ToSingle(xmlNode.SelectSingleNode(AH.A(120493)).InnerText)
			};
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		finally
		{
			XmlNode xmlNode = null;
			xmlDocument = null;
		}
		return result;
	}

	internal void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		try
		{
			A.Tags.Delete(BrandCompliance.TAG_BRAND_COMPLIANCE);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		XmlDocument A2 = new XmlDocument();
		XmlNode xmlNode = A2.CreateElement(BrandCompliance.XML_ROOT);
		A2.AppendChild(xmlNode);
		this.A(ref A2);
		B(ref A2);
		C(ref A2);
		if (xmlNode.HasChildNodes)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A.Tags.Add(BrandCompliance.TAG_BRAND_COMPLIANCE, A2.OuterXml);
		}
		xmlNode = null;
		A2 = null;
	}

	private void A(ref XmlDocument A)
	{
		if (LegalFontTypes != null)
		{
			if (LegalFontTypes.Count > 0)
			{
				goto IL_0068;
			}
			while (true)
			{
				switch (4)
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
		}
		if (!MinFontSize.HasValue)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (!MaxFontSize.HasValue)
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		goto IL_0068;
		IL_0068:
		XmlDocument xmlDocument = A;
		XmlNode xmlNode = xmlDocument.CreateElement(BrandCompliance.XML_LEGAL_FONTS);
		XmlNode xmlNode2 = xmlDocument.CreateElement(BrandCompliance.XML_LEGAL_FONT_TYPES);
		XmlNode newChild = xmlDocument.CreateElement(BrandCompliance.XML_LEGAL_FONT_SIZES);
		XmlNode newChild2 = xmlDocument.CreateElement(BrandCompliance.XML_LEGAL_FONT_COLORS);
		XmlNode xmlNode3 = xmlDocument.CreateElement(BrandCompliance.XML_MIN_FONT_SIZE);
		XmlNode xmlNode4 = xmlDocument.CreateElement(BrandCompliance.XML_MAX_FONT_SIZE);
		XmlNode xmlNode5;
		using (List<string>.Enumerator enumerator = LegalFontTypes.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				xmlNode5 = xmlDocument.CreateElement(AH.A(120502));
				xmlNode5.AppendChild(xmlDocument.CreateCDataSection(current));
				xmlNode2.AppendChild(xmlNode5);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_011f;
				}
				continue;
				end_IL_011f:
				break;
			}
		}
		if (MinFontSize.HasValue)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			xmlNode3.InnerText = Conversions.ToString(MinFontSize.Value);
		}
		if (MaxFontSize.HasValue)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			xmlNode4.InnerText = Conversions.ToString(MaxFontSize.Value);
		}
		XmlNode xmlNode6 = xmlNode;
		xmlNode6.AppendChild(xmlNode2);
		xmlNode6.AppendChild(newChild);
		xmlNode6.AppendChild(newChild2);
		xmlNode6.AppendChild(xmlNode3);
		xmlNode6.AppendChild(xmlNode4);
		_ = null;
		xmlDocument.DocumentElement.AppendChild(xmlNode);
		xmlDocument = null;
		xmlNode = null;
		xmlNode2 = null;
		xmlNode5 = null;
		newChild = null;
		newChild2 = null;
		xmlNode3 = null;
		xmlNode4 = null;
	}

	private void B(ref XmlDocument A)
	{
		if (!TextboxMargins.HasValue)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.A(ref A, TextboxMargins.Value, BrandCompliance.XML_TEXTBOX_MGNS);
			return;
		}
	}

	private void C(ref XmlDocument A)
	{
		if (SlideMargins.HasValue)
		{
			this.A(ref A, SlideMargins.Value, BrandCompliance.XML_SLIDE_MGNS);
		}
	}

	private void A(ref XmlDocument A, Margins B, string C)
	{
		XmlDocument obj = A;
		XmlNode xmlNode = obj.CreateElement(C);
		XmlNode xmlNode2 = obj.CreateElement(AH.A(120462));
		xmlNode2.InnerText = Conversions.ToString(B.Top);
		xmlNode.AppendChild(xmlNode2);
		xmlNode2 = obj.CreateElement(AH.A(120469));
		xmlNode2.InnerText = Conversions.ToString(B.Bottom);
		xmlNode.AppendChild(xmlNode2);
		xmlNode2 = obj.CreateElement(AH.A(120493));
		xmlNode2.InnerText = Conversions.ToString(B.Left);
		xmlNode.AppendChild(xmlNode2);
		xmlNode2 = obj.CreateElement(AH.A(120482));
		xmlNode2.InnerText = Conversions.ToString(B.Right);
		xmlNode.AppendChild(xmlNode2);
		obj.DocumentElement.AppendChild(xmlNode);
		_ = null;
		xmlNode = null;
		xmlNode2 = null;
	}
}
