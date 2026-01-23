using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class ProofingLanguage : BaseTextError
{
	private MsoLanguageID A;

	private MsoLanguageID DefaultLanguageId
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public ProofingLanguage(Range rng, string strSubtitle, MsoLanguageID lang)
		: base(ErrorType.ProofingLanguage, ((Settings)Main.Analysis.Options).ProofingLanguage, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.ProofingLanguage(ref val, strSubtitle);
		DefaultLanguageId = lang;
	}

	public override void FixAction()
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(36293));
		base.Range.LanguageID = (WdLanguageID)DefaultLanguageId;
		undoRecord.EndCustomRecord();
	}
}
