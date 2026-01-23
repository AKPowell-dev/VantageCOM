using System.Drawing;
using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;

namespace Macabacus_Word.Proofing;

public class BaseColorError : BaseError
{
	public BaseColorError(ErrorType errType, Severity sev, object obj, int intColor)
		: base(errType, sev, RuntimeHelpers.GetObjectValue(obj), blnHasFix: false)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).NonconformingColor = ColorTranslator.FromOle(intColor);
		((BaseError)this).HasColorFix = true;
	}
}
