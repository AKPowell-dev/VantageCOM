using System;
using System.Globalization;
using A;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class ProofingLanguage
{
	private MsoLanguageID m_A;

	private MsoLanguageID A
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

	public ProofingLanguage(string strDefaultLanguageId)
	{
		A = (MsoLanguageID)Conversions.ToInteger(strDefaultLanguageId);
	}

	public void Check(Range rng)
	{
		if (rng.LanguageID != (WdLanguageID)A)
		{
			string strSubtitle;
			try
			{
				strSubtitle = XC.A(22739) + CultureInfo.GetCultureInfo((int)A).DisplayName;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				strSubtitle = XC.A(22784);
				ProjectData.ClearProjectError();
			}
			Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.ProofingLanguage(rng.Duplicate, strSubtitle, A));
		}
	}
}
