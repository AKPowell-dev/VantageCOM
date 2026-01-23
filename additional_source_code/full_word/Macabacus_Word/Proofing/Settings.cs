using System;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros.Proofing;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing;

public sealed class Settings : Settings
{
	[CompilerGenerated]
	private int m_A;

	public int MaxFontFamilies
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

	public Settings()
	{
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			((Settings)this).TableCellMargins = A(base.xmlDoc, XC.A(38053));
			MaxFontFamilies = 0;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			((Settings)this).LoadError(ex2);
			ProjectData.ClearProjectError();
		}
		base.xmlDoc = null;
	}

	private Severity A(XmlDocument A, string B)
	{
		return (Severity)Conversions.ToInteger(A.SelectSingleNode(XC.A(37973) + B + XC.A(7149)).Attributes[XC.A(38036)].Value);
	}
}
