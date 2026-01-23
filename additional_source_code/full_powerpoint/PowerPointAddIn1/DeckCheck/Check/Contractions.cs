using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Contractions : BaseTextCheck
{
	[CompilerGenerated]
	private new Dictionary<string, string> A;

	private Dictionary<string, string> Translations
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

	public Contractions()
	{
		base.RegexObj = Text.RegexContractions();
		Translations = Text.ContractionsDictionary();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		if (!strText.Contains(Constants.SINGLE_QUOTE_CLOSE))
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
			if (!strText.Contains(AH.A(15089)))
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
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = base.RegexObj.Matches(strText).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match obj = (Match)enumerator.Current;
				string value = string.Empty;
				Group obj2 = obj.Groups[1];
				if (Translations.TryGetValue(obj2.Value.Replace(Constants.SINGLE_QUOTE_CLOSE, AH.A(15089)).ToLower(), out value))
				{
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.Contractions(sld, shp, para.get_Characters(checked(obj2.Index + 1), obj2.Length), value));
				}
				obj2 = null;
			}
			while (true)
			{
				switch (1)
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
					switch (2)
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
}
