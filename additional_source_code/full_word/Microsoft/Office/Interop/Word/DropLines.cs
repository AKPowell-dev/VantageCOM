using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[Guid("9F1DF642-3CCE-4D83-A770-D2634A05D278")]
[CompilerGenerated]
public interface DropLines
{
	void _VtblGap1_5();

	[DispId(1610743813)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743813)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
