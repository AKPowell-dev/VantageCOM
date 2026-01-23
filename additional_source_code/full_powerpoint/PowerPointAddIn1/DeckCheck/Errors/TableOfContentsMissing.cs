using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TableOfContentsMissing : BaseError
{
	public TableOfContentsMissing()
		: base(ErrorType.AgendaMissing, Main.Analysis.Options.CheckAgendaUpdated, null, null, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(18869);
		((BaseError)this).Subtitle = AH.A(18920);
	}
}
