using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Agenda;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AgendaNotUpdated : BaseError
{
	public AgendaNotUpdated()
		: base(ErrorType.AgendaNotUpdated, Main.Analysis.Options.CheckAgendaUpdated, null, null, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(18696);
		((BaseError)this).Subtitle = AH.A(18733);
	}

	public override void FixAction()
	{
		Application application = NG.A.Application;
		application.StartNewUndoEntry();
		Update.A(A: false, application.ActivePresentation);
		application = null;
	}
}
