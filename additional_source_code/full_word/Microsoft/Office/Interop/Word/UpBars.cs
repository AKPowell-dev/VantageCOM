using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("86905AC9-33F3-4A88-96C8-B289B0390BCA")]
public interface UpBars
{
	void _VtblGap1_7();

	[DispId(1610743815)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743815)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
