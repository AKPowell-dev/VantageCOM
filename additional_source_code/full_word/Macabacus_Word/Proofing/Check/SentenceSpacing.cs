using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class SentenceSpacing : BaseTextCheck
{
	public SentenceSpacing(SpacesBetweenSentences conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Invalid comparison between Unknown and I4
		if ((int)conv == 1)
		{
			base.RegexObj = new Regex(XC.A(25309));
			base.Fix = XC.A(25376);
		}
		else
		{
			base.RegexObj = new Regex(XC.A(25383));
			base.Fix = XC.A(21362);
		}
	}

	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(XC.A(4860)))
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
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
				MatchCollection matchCollection = base.RegexObj.Matches(strText);
				enumerator = matchCollection.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Match obj = (Match)enumerator.Current;
						Range duplicate = rng.Duplicate;
						Group obj2 = obj.Groups[1];
						duplicate.SetRange(rng.Characters[obj2.Index + 1].Start, rng.Characters[obj2.Index + obj2.Length].End);
						obj2 = null;
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.SentenceSpacing(duplicate, base.Fix));
						duplicate = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_00ec;
						}
						continue;
						end_IL_00ec:
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
				matchCollection = null;
				return;
			}
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(XC.A(4860)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			MatchCollection matchCollection = base.RegexObj.Matches(strText);
			enumerator = matchCollection.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.SentenceSpacing(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), base.Fix));
					obj = null;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00ab;
					}
					continue;
					end_IL_00ab:
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
			matchCollection = null;
			return;
		}
	}
}
