using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class BulletSize : BaseTextError
{
	[CompilerGenerated]
	private new List<float> A;

	private List<float> FixOptions
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

	public BulletSize(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<string> listLabels, string strSubtitle, List<TextRange2> listRanges, List<float> listFixes)
		: base(ErrorType.BulletSize, Main.Analysis.Options.BulletSize, sld, shp, listRanges, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = AH.A(42470);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(42519);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		float relativeSize = FixOptions[i];
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			textRange.ParagraphFormat.Bullet.RelativeSize = relativeSize;
		}
	}
}
