using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class BulletIndentation : BaseTextError
{
	[CompilerGenerated]
	private new List<BulletIndentFix> A;

	private List<BulletIndentFix> FixOptions
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

	public BulletIndentation(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<string> listLabels, string strSubtitle, List<TextRange2> listRanges, List<BulletIndentFix> listFixes)
		: base(ErrorType.BulletIndent, Main.Analysis.Options.BulletIndentation, sld, shp, listRanges, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = AH.A(41877);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(41940);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		BulletIndentFix bulletIndentFix = FixOptions[i];
		float leftIndent = bulletIndentFix.LeftIndent;
		float firstLineIndent = bulletIndentFix.FirstLineIndent;
		using IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator();
		while (enumerator.MoveNext())
		{
			ParagraphFormat2 paragraphFormat = enumerator.Current.ParagraphFormat;
			paragraphFormat.LeftIndent = leftIndent;
			paragraphFormat.FirstLineIndent = firstLineIndent;
			_ = null;
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
			return;
		}
	}
}
