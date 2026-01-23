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

public sealed class DashSpacingInconsistent : BaseTextCheck
{
	public DashSpacingInconsistent(DashSpacing conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Invalid comparison between Unknown and I4
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
					base.RegexObj = new Regex(XC.A(24592));
					base.Fix = XC.A(24589);
					return;
				}
			}
		}
		base.RegexObj = new Regex(XC.A(24607));
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
				switch (3)
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
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.DashSpacingInconsistent(duplicate, base.Fix));
							duplicate = null;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_00e2;
							}
							continue;
							end_IL_00e2:
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
					Group obj = ((Match)enumerator.Current).Groups[1];
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.DashSpacingInconsistent(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), base.Fix));
					obj = null;
				}
				while (true)
				{
					switch (3)
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
