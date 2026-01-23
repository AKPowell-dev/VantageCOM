using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("84A6A663-AEF4-4FCD-83FD-9BB707F157CA")]
[TypeIdentifier]
public interface DownBars
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
