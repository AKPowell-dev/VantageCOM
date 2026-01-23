using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[Guid("00020961-0000-0000-C000-000000000046")]
[TypeIdentifier]
[CompilerGenerated]
public interface Windows : IEnumerable
{
	void _VtblGap1_1();

	[DispId(2)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		get;
	}

	void _VtblGap2_3();

	[DispId(0)]
	Window this[[In][MarshalAs(UnmanagedType.Struct)] ref object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
