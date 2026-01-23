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

public sealed class MissingParentheses : BaseTextCheck
{
	private Regex m_A;

	private Regex m_B;

	private Regex A
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

	private Regex B
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
		}
	}

	public MissingParentheses()
	{
		base.RegexObj = new Regex(XC.A(25274));
		A = new Regex(XC.A(25299));
		B = new Regex(XC.A(25304));
	}

	public override void Check(Range rng, string strText)
	{
		string text = strText;
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = matchCollection.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				text = Text.MaskText(text, match);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		matchCollection = A.Matches(text);
		IEnumerator enumerator2 = matchCollection.GetEnumerator();
		checked
		{
			try
			{
				while (enumerator2.MoveNext())
				{
					Match match2 = (Match)enumerator2.Current;
					Range duplicate = rng.Duplicate;
					duplicate.SetRange(rng.Characters[match2.Index + 1].Start, rng.Characters[match2.Index + match2.Length].End);
					Main.Analysis.Errors.Add(new MissingClosingParenthesis(duplicate));
					duplicate = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_010f;
					}
					continue;
					end_IL_010f:
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator2 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			matchCollection = B.Matches(text);
			IEnumerator enumerator3 = default(IEnumerator);
			try
			{
				enumerator3 = matchCollection.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					Match match3 = (Match)enumerator3.Current;
					Range duplicate2 = rng.Duplicate;
					duplicate2.SetRange(rng.Characters[match3.Index + 1].Start, rng.Characters[match3.Index + match3.Length].End);
					Main.Analysis.Errors.Add(new MissingOpeningParenthesis(duplicate2));
					duplicate2 = null;
				}
			}
			finally
			{
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator3 as IDisposable).Dispose();
						break;
					}
				}
			}
			matchCollection = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		string text = strText;
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = matchCollection.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				text = Text.MaskText(text, match);
			}
			while (true)
			{
				switch (1)
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
		matchCollection = A.Matches(text);
		checked
		{
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = matchCollection.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Match match2 = (Match)enumerator2.Current;
					Main.Analysis.Errors.Add(new MissingClosingParenthesis(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(match2.Index + 1, match2.Length)));
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_00d3;
					}
					continue;
					end_IL_00d3:
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
			matchCollection = B.Matches(text);
			IEnumerator enumerator3 = matchCollection.GetEnumerator();
			try
			{
				while (enumerator3.MoveNext())
				{
					Match match3 = (Match)enumerator3.Current;
					Main.Analysis.Errors.Add(new MissingOpeningParenthesis(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(match3.Index + 1, match3.Length)));
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_016f;
					}
					continue;
					end_IL_016f:
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator3 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			matchCollection = null;
		}
	}
}
