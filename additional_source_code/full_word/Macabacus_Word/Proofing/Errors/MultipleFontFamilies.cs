using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class MultipleFontFamilies : BaseTextError
{
	private List<string> m_A;

	private List<string> FixOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public MultipleFontFamilies(Range rng, List<string> listLabels, string strSubtitle, List<string> listFixes)
		: base(ErrorType.MultipleFontFamilies, ((Settings)Main.Analysis.Options).MultipleFontFamilies, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		A(listLabels, strSubtitle, listFixes);
	}

	public MultipleFontFamilies(object shp, TextRange2 rng, List<string> listLabels, string strSubtitle, List<string> listFixes)
		: base(ErrorType.MultipleFontFamilies, ((Settings)Main.Analysis.Options).MultipleFontFamilies, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		A(listLabels, strSubtitle, listFixes);
	}

	private void A(List<string> A, string B, List<string> C)
	{
		((BaseError)this).DisplayText = A;
		FixOptions = C;
		((BaseError)this).Title = XC.A(35943);
		((BaseError)this).Subtitle = B;
		((BaseError)this).Tooltip = XC.A(35988);
	}

	public override void FixAction(int i)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(36258));
		string name = FixOptions[i];
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			textRange.Font.Name = name;
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
