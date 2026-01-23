using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlaceholderFontColorMismatch : BaseError
{
	[CompilerGenerated]
	private new int A;

	private int MasterColor
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

	public PlaceholderFontColorMismatch(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, int intColorMaster)
		: base(ErrorType.PlaceholderFontColorMismatch, ((Settings)Main.Analysis.Options).ColorPalette, sld, shp, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		ColorTranslator.FromOle(intColorMaster);
		((BaseError)this).Title = AH.A(38679);
		((BaseError)this).Subtitle = AH.A(38742);
		((BaseError)this).TextRanges = listRanges;
		MasterColor = intColorMaster;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		using IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator();
		while (enumerator.MoveNext())
		{
			enumerator.Current.Font.Fill.ForeColor.RGB = MasterColor;
		}
		while (true)
		{
			switch (2)
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
