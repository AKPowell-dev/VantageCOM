using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class HyphenSpacingImproper : BaseTextCheck
{
	private List<string> m_A;

	private List<string> m_B;

	private List<string> A
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

	private List<string> B
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

	public HyphenSpacingImproper()
	{
		base.RegexObj = new Regex(XC.A(24925));
		B = new List<string>(new string[3]
		{
			XC.A(24942),
			XC.A(24997),
			XC.A(25056)
		});
		A = new List<string>(new string[3]
		{
			XC.A(6388),
			XC.A(24622),
			XC.A(24589)
		});
	}

	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(XC.A(6388)))
		{
			return;
		}
		checked
		{
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
				MatchCollection matchCollection;
				try
				{
					matchCollection = base.RegexObj.Matches(strText);
					try
					{
						enumerator = matchCollection.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Match match = (Match)enumerator.Current;
							if (match.Length > 1)
							{
								Range duplicate = rng.Duplicate;
								Match match2 = match;
								duplicate.SetRange(rng.Characters[match2.Index + 1].Start, rng.Characters[match2.Index + match2.Length].End);
								match2 = null;
								Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.HyphenSpacingImproper(duplicate, B, A));
								duplicate = null;
							}
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_00e9;
							}
							continue;
							end_IL_00e9:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (1)
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
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				matchCollection = null;
				return;
			}
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(XC.A(6388)))
		{
			return;
		}
		MatchCollection matchCollection;
		try
		{
			matchCollection = base.RegexObj.Matches(strText);
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = matchCollection.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Match match = (Match)enumerator.Current;
					if (match.Length <= 1)
					{
						continue;
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
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.HyphenSpacingImproper(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(match.Index + 1), match.Length), B, A));
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_00b1;
					}
					continue;
					end_IL_00b1:
					break;
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		matchCollection = null;
	}
}
