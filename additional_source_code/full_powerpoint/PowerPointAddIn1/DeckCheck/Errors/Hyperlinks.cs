using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Links;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class Hyperlinks : BaseError
{
	public Hyperlinks(Slide sld, int i)
		: base(ErrorType.Hyperlinks, Main.Analysis.Options.Hyperlinks, sld, null, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(36348);
		((BaseError)this).Subtitle = AH.A(36272) + sld.SlideIndex + AH.A(36369) + i + AH.A(36390);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		Microsoft.Office.Interop.PowerPoint.Hyperlinks hyperlinks = base.Slide.Hyperlinks;
		for (int i = hyperlinks.Count; i >= 1; i = checked(i + -1))
		{
			PowerPointAddIn1.Links.Hyperlinks.C(hyperlinks[i]);
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			hyperlinks = null;
			return;
		}
	}
}
