using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class RepeatedWords : BaseTextError
{
	public RepeatedWords(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strSubtitle)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).RepeatedWords, sld, shp, rng, blnHasFix: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.RepeatedWords(ref val, strSubtitle);
	}

	public override void FixAction(int i)
	{
		Regex regex = new Regex(AH.A(46451), RegexOptions.IgnoreCase);
		NG.A.Application.StartNewUndoEntry();
		TextRange2 textRange = ((BaseError)this).TextRanges[0];
		MatchCollection matchCollection = regex.Matches(textRange.Text);
		IEnumerator enumerator = default(IEnumerator);
		while (matchCollection.Count > 0)
		{
			{
				enumerator = matchCollection.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Match match = (Match)enumerator.Current;
						textRange.Text = ((BaseError)this).TextRanges[0].get_Characters(1, match.Groups[1].Length).Text;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						break;
					}
				}
				finally
				{
					IDisposable disposable = enumerator as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
			}
			matchCollection = regex.Matches(textRange.Text);
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			textRange = null;
			regex = null;
			matchCollection = null;
			return;
		}
	}
}
