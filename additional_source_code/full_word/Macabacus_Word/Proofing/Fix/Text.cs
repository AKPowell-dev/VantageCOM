using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Fix;

public sealed class Text
{
	public static void ReplaceText(BaseError err, int i)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(36336));
		if ((object)((object)err).GetType() == typeof(DoubleQuotesStyle))
		{
			while (true)
			{
				switch (6)
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
			if (Operators.CompareString(((BaseError)err).ReplacementText[i], XC.A(24629), TextCompare: false) != 0)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (((BaseError)err).TextRanges[0].Start == 1)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					((BaseError)err).TextRanges[0].Text = Constants.DOUBLE_QUOTE_OPEN;
				}
				else if (Operators.CompareString(err.Shape.TextFrame2.TextRange.get_Characters(checked(((BaseError)err).TextRanges[0].Start - 1), 1).Text, XC.A(18458), TextCompare: false) == 0)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					((BaseError)err).TextRanges[0].Text = Constants.DOUBLE_QUOTE_OPEN;
				}
				else
				{
					((BaseError)err).TextRanges[0].Text = Constants.DOUBLE_QUOTE_CLOSE;
				}
				goto IL_0172;
			}
		}
		err.Ranges[0].Text = ((BaseError)err).ReplacementText[i];
		goto IL_0172;
		IL_0172:
		undoRecord.EndCustomRecord();
	}
}
