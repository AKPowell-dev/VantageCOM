using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SpellingAdviser : BaseTextError
{
	[CompilerGenerated]
	private new AdviserSpelling A;

	private AdviserSpelling Convention
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

	public SpellingAdviser(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, AdviserSpelling conv)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).SpellingAdviser, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cc: Unknown result type (might be due to invalid IL or missing references)
		string text;
		if (listRanges.Count == 1)
		{
			text = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
		}
		else if ((int)conv == 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			text = AH.A(42825) + listRanges.Count + AH.A(46826);
		}
		else
		{
			text = AH.A(42825) + listRanges.Count + AH.A(46877);
		}
		BaseError val = (BaseError)(object)this;
		Errors.SpellingAdviser(ref val, text);
		Convention = conv;
	}

	public override void FixAction(int i)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Unknown result type (might be due to invalid IL or missing references)
		NG.A.Application.StartNewUndoEntry();
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			if ((int)Convention == 0)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				textRange.Text = textRange.Text.Replace(AH.A(46798), AH.A(46805));
				textRange.Text = textRange.Text.Replace(AH.A(46812), AH.A(46819));
			}
			else
			{
				textRange.Text = textRange.Text.Replace(AH.A(46805), AH.A(46798));
				textRange.Text = textRange.Text.Replace(AH.A(46819), AH.A(46812));
			}
			TextRange2 current = null;
		}
	}
}
