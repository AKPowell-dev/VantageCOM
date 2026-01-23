using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlaceholderFontStyleMismatch : BaseError
{
	[CompilerGenerated]
	private new Font2 A;

	private Font2 MasterFont
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

	public PlaceholderFontStyleMismatch(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, int intLevel, Font2 font)
		: base(ErrorType.PlaceholderFontStyleMismatch, Main.Analysis.Options.CheckPlaceholderFontStyleMismatch, sld, shp, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(38906);
		((BaseError)this).Subtitle = AH.A(38015) + intLevel + AH.A(38969) + font.Name + AH.A(39115) + string.Format(AH.A(39130), font.Size) + AH.A(14255);
		((BaseError)this).Tooltip = AH.A(38015) + intLevel + AH.A(39149);
		((BaseError)this).TextRanges = listRanges;
		MasterFont = font;
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
				Font2 font = enumerator.Current.Font;
				font.Name = MasterFont.Name;
				font.Size = MasterFont.Size;
				_ = null;
			}
			while (true)
			{
				switch (5)
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
					switch (6)
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
