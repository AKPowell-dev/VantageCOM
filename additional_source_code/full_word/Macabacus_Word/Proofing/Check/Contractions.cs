using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class Contractions : BaseTextCheck
{
	private Dictionary<string, string> m_A;

	private Dictionary<string, string> A
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

	public Contractions()
	{
		base.RegexObj = Text.RegexContractions();
		A = Text.ContractionsDictionary();
	}

	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(Constants.SINGLE_QUOTE_CLOSE))
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
			if (!strText.Contains(XC.A(6376)))
			{
				goto IL_0158;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = matchCollection.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Match obj = (Match)enumerator.Current;
					string value = string.Empty;
					Range duplicate = rng.Duplicate;
					Group obj2 = obj.Groups[1];
					duplicate.SetRange(rng.Characters[obj2.Index + 1].Start, rng.Characters[obj2.Index + obj2.Length].End);
					A.TryGetValue(obj2.Value.Replace(Constants.SINGLE_QUOTE_CLOSE, XC.A(6376)), out value);
					obj2 = null;
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.Contractions(duplicate, value));
					duplicate = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_012e;
					}
					continue;
					end_IL_012e:
					break;
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
			goto IL_0158;
		}
		IL_0158:
		matchCollection = null;
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(Constants.SINGLE_QUOTE_CLOSE))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!strText.Contains(XC.A(6376)))
			{
				goto IL_0121;
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
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = matchCollection.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match obj = (Match)enumerator.Current;
				string value = string.Empty;
				Group obj2 = obj.Groups[1];
				A.TryGetValue(obj2.Value.Replace(Constants.SINGLE_QUOTE_CLOSE, XC.A(6376)), out value);
				Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.Contractions(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj2.Index + 1), obj2.Length), value));
				obj2 = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_00f7;
				}
				continue;
				end_IL_00f7:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		goto IL_0121;
		IL_0121:
		matchCollection = null;
	}
}
