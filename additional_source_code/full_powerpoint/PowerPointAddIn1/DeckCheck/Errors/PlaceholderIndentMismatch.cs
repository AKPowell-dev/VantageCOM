using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlaceholderIndentMismatch : BaseError
{
	[CompilerGenerated]
	private new float A;

	[CompilerGenerated]
	private float B;

	[CompilerGenerated]
	private float C;

	private float FirstLineIndent
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

	private float LeftIndent
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	private float RightIndent
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public PlaceholderIndentMismatch(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, int intLevel, float sngFirstLineIndent, float sngLeftIndent, float sngRightIndent)
		: base(ErrorType.PlaceholderIndentMismatch, Main.Analysis.Options.CheckPlaceholderIndentMismatch, sld, shp, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(39355);
		((BaseError)this).Subtitle = AH.A(38015) + intLevel + AH.A(39410) + string.Format(AH.A(39533), sngFirstLineIndent, sngLeftIndent, sngRightIndent) + AH.A(14255);
		((BaseError)this).Tooltip = AH.A(38015) + intLevel + AH.A(39584) + string.Format(AH.A(39760), sngFirstLineIndent, sngLeftIndent, sngRightIndent) + AH.A(39867);
		((BaseError)this).TextRanges = listRanges;
		FirstLineIndent = sngFirstLineIndent;
		LeftIndent = sngLeftIndent;
		RightIndent = sngRightIndent;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ParagraphFormat2 paragraphFormat = enumerator.Current.ParagraphFormat;
				paragraphFormat.LeftIndent = LeftIndent;
				paragraphFormat.RightIndent = RightIndent;
				paragraphFormat.FirstLineIndent = FirstLineIndent;
				_ = null;
			}
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
				return;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (2)
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
