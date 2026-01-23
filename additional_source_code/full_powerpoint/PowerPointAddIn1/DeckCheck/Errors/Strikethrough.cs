using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class Strikethrough : BaseTextError
{
	public Strikethrough(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.StrikethroughFont, ((Settings)Main.Analysis.Options).StrikethroughFont, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(18444);
		((BaseError)this).Subtitle = AH.A(18481);
		((BaseError)this).Tooltip = AH.A(18548);
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
				TextRange2 current = enumerator.Current;
				current.Font.StrikeThrough = MsoTriState.msoFalse;
				current.Font.DoubleStrikeThrough = MsoTriState.msoFalse;
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
