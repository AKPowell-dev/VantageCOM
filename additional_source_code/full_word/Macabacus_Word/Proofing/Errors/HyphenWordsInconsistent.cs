using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Errors;

public sealed class HyphenWordsInconsistent : BaseTextError
{
	public HyphenWordsInconsistent(Range rng, List<string> listLabels, List<string> listFixes)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).HyphenWordsInconsistent, rng, blnHasFix: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.HyphenWordsInconsistent(ref val, GenerateSnippet(rng), listLabels, listFixes);
	}

	public HyphenWordsInconsistent(object shp, TextRange2 rng, List<string> listLabels, List<string> listFixes)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).HyphenWordsInconsistent, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.HyphenWordsInconsistent(ref val, GenerateSnippet(rng), listLabels, listFixes);
	}

	public override void FixAction(int i)
	{
		string text = ((BaseError)this).ReplacementText[i];
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(35523));
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				if (Operators.CompareString(Strings.Left(current.Text, 1), Strings.Left(current.Text.ToUpper(), 1), TextCompare: false) == 0)
				{
					while (true)
					{
						switch (7)
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
					text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text);
				}
				current.Text = text;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (7)
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
