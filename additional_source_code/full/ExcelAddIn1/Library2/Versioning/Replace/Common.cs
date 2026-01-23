using MacabacusMacros;
using MacabacusMacros.Libraries.Versioning;

namespace ExcelAddIn1.Library2.Versioning.Replace;

public sealed class Common
{
	internal static string A(ContentItem A)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		return CloudStorage.FillPlaceholdersInPath(A.ContentInfo.ContentPath);
	}
}
