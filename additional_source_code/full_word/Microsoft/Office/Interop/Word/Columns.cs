using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[Guid("0002094B-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface Columns : IEnumerable
{
	void _VtblGap1_1();

	[DispId(2)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		get;
	}

	void _VtblGap2_10();

	[DispId(0)]
	Column this[[In] int Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(5)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Column Add([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object BeforeColumn);
}
