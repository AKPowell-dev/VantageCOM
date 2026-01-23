using System.Runtime.CompilerServices;
using System.Xml;
using A;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Template;

public sealed class Options
{
	public enum RuleViolationEnum
	{
		WarnOnSave,
		BlockSaveAndWarn
	}

	public enum LegalCheckActionEnum
	{
		DoNotCheck,
		CheckAgainstOriginalTemplate,
		CheckAgainstAllTemplates
	}

	public enum LegalCheckScopeEnum
	{
		Restricted,
		Unrestricted,
		Both
	}

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private LegalCheckActionEnum A;

	[CompilerGenerated]
	private LegalCheckScopeEnum A;

	[CompilerGenerated]
	private RuleViolationEnum A;

	public bool RequireLegalSlide
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

	public bool RequireContactSlide
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public bool RequireFrontCover
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public bool RequireBackCover
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	public LegalCheckActionEnum LegalCheckAction
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

	public LegalCheckScopeEnum LegalCheckScope
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

	public RuleViolationEnum RuleViolationAction
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

	public Options(XmlDocument xmlDoc)
	{
		string text = AH.A(137172);
		XmlDocument xmlDocument = xmlDoc;
		RequireLegalSlide = Conversions.ToBoolean(xmlDocument.SelectSingleNode(AH.A(93759) + text + AH.A(137199)).InnerText);
		RequireContactSlide = Conversions.ToBoolean(xmlDocument.SelectSingleNode(AH.A(93759) + text + AH.A(137242)).InnerText);
		RequireFrontCover = Conversions.ToBoolean(xmlDocument.SelectSingleNode(AH.A(93759) + text + AH.A(137289)).InnerText);
		RequireBackCover = Conversions.ToBoolean(xmlDocument.SelectSingleNode(AH.A(93759) + text + AH.A(137342)).InnerText);
		LegalCheckAction = (LegalCheckActionEnum)Conversions.ToInteger(xmlDocument.SelectSingleNode(AH.A(93759) + text + AH.A(137393)).InnerText);
		LegalCheckScope = (LegalCheckScopeEnum)Conversions.ToInteger(xmlDocument.SelectSingleNode(AH.A(93759) + text + AH.A(137434)).InnerText);
		RuleViolationAction = (RuleViolationEnum)Conversions.ToInteger(xmlDocument.SelectSingleNode(AH.A(93759) + text + AH.A(137473)).InnerText);
		xmlDocument = null;
	}
}
