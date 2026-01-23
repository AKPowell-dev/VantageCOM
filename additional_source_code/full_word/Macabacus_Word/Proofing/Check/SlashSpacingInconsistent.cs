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

public sealed class SlashSpacingInconsistent : BaseTextCheck
{
	public SlashSpacingInconsistent(SlashSpacing conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Invalid comparison between Unknown and I4
		if ((int)conv == 1)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.RegexObj = new Regex(XC.A(25453));
					base.Fix = XC.A(25450);
					return;
				}
			}
		}
		base.RegexObj = new Regex(XC.A(25468));
		base.Fix = XC.A(25483);
	}

	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(XC.A(25450)))
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
				MatchCollection matchCollection;
				try
				{
					matchCollection = base.RegexObj.Matches(strText);
					try
					{
						enumerator = matchCollection.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Match obj = (Match)enumerator.Current;
							Range duplicate = rng.Duplicate;
							Group obj2 = obj.Groups[1];
							duplicate.SetRange(rng.Characters[obj2.Index + 1].Start, rng.Characters[obj2.Index + obj2.Length].End);
							obj2 = null;
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.SlashSpacingInconsistent(duplicate, base.Fix));
							duplicate = null;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_00ea;
							}
							continue;
							end_IL_00ea:
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
		if (!strText.Contains(XC.A(25450)))
		{
			return;
		}
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
						Group obj = ((Match)enumerator.Current).Groups[1];
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.SlashSpacingInconsistent(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), base.Fix));
						obj = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_00a5;
						}
						continue;
						end_IL_00a5:
						break;
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
