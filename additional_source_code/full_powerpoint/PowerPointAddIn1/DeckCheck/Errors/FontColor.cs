using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FontColor : BaseColorError
{
	public FontColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, IList<TextRange2> listRanges, Severity sev)
		: base(ErrorType.ColorPaletteFont, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(26171);
		((BaseError)this).Subtitle = AH.A(26192);
		((BaseError)this).TextRanges = listRanges;
	}

	public override void FixAction(Color color)
	{
		int rGB = ColorTranslator.ToOle(color);
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.Font.Fill.ForeColor.RGB = rGB;
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
