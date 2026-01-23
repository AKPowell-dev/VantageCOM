using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class DashSpacingUnbalanced : BaseTextCheck
{
	public DashSpacingUnbalanced(DashSpacing conv)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Invalid comparison between Unknown and I4
		base.RegexObj = new Regex(Constants.REGEX_DASH_SPACING);
		if ((int)conv == 1)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.Fix = XC.A(24589);
					return;
				}
			}
		}
		base.Fix = XC.A(24622);
	}

	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(XC.A(24589)))
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
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
							if (match.Groups[0].Length == 2)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								Range duplicate = rng.Duplicate;
								duplicate.SetRange(rng.Characters[match.Index + 1].Start, rng.Characters[match.Index + match.Length].End);
								Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.DashSpacingUnbalanced(duplicate, base.Fix));
								duplicate = null;
							}
							_ = null;
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
				return;
			}
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(XC.A(24589)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (7)
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
						Group obj = ((Match)enumerator.Current).Groups[0];
						if (obj.Length == 2)
						{
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.DashSpacingUnbalanced(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), base.Fix));
						}
						obj = null;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_00b6;
						}
						continue;
						end_IL_00b6:
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
