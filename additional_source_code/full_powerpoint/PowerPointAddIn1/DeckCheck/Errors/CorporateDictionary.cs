using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.CorporateDictionary;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class CorporateDictionary : BaseTextError
{
	[CompilerGenerated]
	private new Rule A;

	[CompilerGenerated]
	private new string A;

	private Rule Rule
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal string RuleId
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

	public CorporateDictionary(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, Rule _rule)
		: base(ErrorType.CorporateDictionary, _rule.Severity, sld, shp, listRanges, _rule.ReplaceWith.Count > 0, _rule.ReplaceWith.Count > 0)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		Rule = _rule;
		RuleId = _rule.Id;
		BaseError val = (BaseError)(object)this;
		Errors.CorporateDictionary(ref val, _rule);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				current.Text = PowerPointAddIn1.DeckCheck.Fix.Text.A(current.Text, ((BaseError)this).ReplacementText[i], Rule);
			}
		}
		finally
		{
			if (enumerator != null)
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
					enumerator.Dispose();
					break;
				}
			}
		}
	}
}
