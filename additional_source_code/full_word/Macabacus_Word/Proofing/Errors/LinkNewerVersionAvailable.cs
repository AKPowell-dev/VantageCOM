using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;

namespace Macabacus_Word.Proofing.Errors;

public sealed class LinkNewerVersionAvailable : BaseError
{
	public LinkNewerVersionAvailable(object obj, string strSubtitle)
		: base(ErrorType.LinkNewerVersionAvailable, ((Settings)Main.Analysis.Options).CheckLinks, RuntimeHelpers.GetObjectValue(obj), blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = XC.A(26577);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = XC.A(26632);
	}
}
