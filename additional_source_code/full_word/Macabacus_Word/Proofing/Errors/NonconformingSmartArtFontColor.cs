using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingSmartArtFontColor : BaseColorError
{
	public NonconformingSmartArtFontColor(object shp, int intColor, List<TextRange2> listShapes, Severity sev)
		: base(ErrorType.ColorPaletteFont, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).TextRanges = listShapes;
		((BaseError)this).Title = XC.A(32519);
		((BaseError)this).Subtitle = XC.A(32570);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		int rGB = ColorTranslator.ToOle(color);
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
				switch (3)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
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
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
