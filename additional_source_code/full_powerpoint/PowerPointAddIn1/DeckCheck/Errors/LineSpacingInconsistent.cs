using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LineSpacingInconsistent : BaseTextError
{
	[CompilerGenerated]
	private new List<LineSpacingFix> A;

	private List<LineSpacingFix> FixOptions
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

	public LineSpacingInconsistent(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<string> listLabels, string strSubtitle, List<TextRange2> listRanges, List<LineSpacingFix> listFixes)
		: base(ErrorType.LineSpacing, Main.Analysis.Options.LineSpacing, sld, shp, listRanges, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = AH.A(45435);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(45486);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		LineSpacingFix lineSpacingFix = FixOptions[i];
		float spaceBefore = lineSpacingFix.SpaceBefore;
		float spaceAfter = lineSpacingFix.SpaceAfter;
		float spaceWithin = lineSpacingFix.SpaceWithin;
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ParagraphFormat2 paragraphFormat = enumerator.Current.ParagraphFormat;
				paragraphFormat.SpaceBefore = spaceBefore;
				paragraphFormat.SpaceAfter = spaceAfter;
				paragraphFormat.SpaceWithin = spaceWithin;
				_ = null;
			}
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
				return;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}
}
