using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TextOutlineColor : BaseColorError
{
	public TextOutlineColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, IList<TextRange2> listRanges, Severity sev)
		: base(ErrorType.ColorPaletteFont, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(29064);
		((BaseError)this).Subtitle = AH.A(29101);
		((BaseError)this).TextRanges = listRanges;
	}

	public override void FixAction(Color color)
	{
		bool flag = false;
		NG.A.Application.StartNewUndoEntry();
		using (IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				int num = ColorTranslator.ToOle(color);
				Microsoft.Office.Core.LineFormat line = current.Font.Line;
				line.ForeColor.RGB = num;
				if (line.ForeColor.RGB != num)
				{
					while (true)
					{
						switch (1)
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
					flag = true;
				}
				_ = null;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0086;
				}
				continue;
				end_IL_0086:
				break;
			}
		}
		if (!flag || base.Shape.HasTable != MsoTriState.msoTrue)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			Forms.ErrorMessage(AH.A(28880));
			return;
		}
	}
}
