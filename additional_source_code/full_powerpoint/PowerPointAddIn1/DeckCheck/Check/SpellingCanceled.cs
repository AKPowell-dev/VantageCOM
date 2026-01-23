using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class SpellingCanceled : BaseTextCheck
{
	[CompilerGenerated]
	private new CanceledSpelling A;

	private CanceledSpelling Convention
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return A;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			A = value;
		}
	}

	public SpellingCanceled(CanceledSpelling conv)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		Convention = conv;
		if ((int)conv == 0)
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
					base.RegexObj = Text.RegexCancelled();
					return;
				}
			}
		}
		base.RegexObj = Text.RegexCanceled();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		//IL_00b4: Unknown result type (might be due to invalid IL or missing references)
		List<TextRange2> list = new List<TextRange2>();
		IEnumerator enumerator = base.RegexObj.Matches(strText).GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Group obj = ((Match)enumerator.Current).Groups[1];
				list.Add(para.get_Characters(checked(obj.Index + 1), obj.Length));
				obj = null;
			}
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
		if (list.Count > 0)
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
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.SpellingCanceled(sld, shp, list, Convention));
		}
		list = null;
	}
}
