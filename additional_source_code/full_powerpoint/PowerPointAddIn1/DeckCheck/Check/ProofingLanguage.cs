using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ProofingLanguage : BaseCheck
{
	[CompilerGenerated]
	private MsoLanguageID A;

	[CompilerGenerated]
	private List<TextRange2> A;

	private MsoLanguageID DefaultLanguageId
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

	private List<TextRange2> TextRanges
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

	public ProofingLanguage(string strDefaultLanguageId)
	{
		DefaultLanguageId = (MsoLanguageID)Conversions.ToInteger(strDefaultLanguageId);
		TextRanges = new List<TextRange2>();
	}

	public void CheckParagraph(TextRange2 para)
	{
		if (para.LanguageID != DefaultLanguageId)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					TextRanges.Add(para);
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = para.get_Runs(-1, -1).GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 textRange = (TextRange2)enumerator.Current;
				if (textRange.LanguageID != DefaultLanguageId)
				{
					TextRanges.Add(textRange);
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (TextRanges.Count <= 0)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			string strSubtitle;
			try
			{
				strSubtitle = AH.A(14653) + CultureInfo.GetCultureInfo((int)DefaultLanguageId).DisplayName;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				strSubtitle = AH.A(14698);
				ProjectData.ClearProjectError();
			}
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.ProofingLanguage(sld, shp, TextRanges.ToList(), strSubtitle, DefaultLanguageId));
			TextRanges.Clear();
			return;
		}
	}
}
