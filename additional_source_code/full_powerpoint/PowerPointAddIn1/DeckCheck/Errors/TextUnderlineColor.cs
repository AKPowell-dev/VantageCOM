using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TextUnderlineColor : BaseColorError
{
	public TextUnderlineColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, IList<TextRange2> listRanges, Severity sev)
		: base(ErrorType.ColorPaletteFont, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(29226);
		((BaseError)this).Subtitle = AH.A(29267);
		((BaseError)this).TextRanges = listRanges;
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Font2 font = enumerator.Current.Font;
				float size = font.Size;
				string name = font.Name;
				int rGB = font.Fill.ForeColor.RGB;
				font.UnderlineColor.RGB = ColorTranslator.ToOle(color);
				font.Size = size;
				font.Name = name;
				font.Fill.ForeColor.RGB = rGB;
				_ = null;
			}
			while (true)
			{
				switch (3)
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
