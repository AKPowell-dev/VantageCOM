using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("00020940-0000-0000-C000-000000000046")]
public interface Comments : IEnumerable
{
	void _VtblGap1_1();

	[DispId(2)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(2)]
		get;
	}

	void _VtblGap2_5();

	[DispId(0)]
	Comment this[[In] int Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
