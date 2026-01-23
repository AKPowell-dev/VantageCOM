using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[Guid("000C0302-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface _CommandBars : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_11();

	[DispId(0)]
	CommandBar this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_19();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809373)]
	void ExecuteMso([In][MarshalAs(UnmanagedType.BStr)] string idMso);

	void _VtblGap3_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1610809376)]
	bool GetPressedMso([In][MarshalAs(UnmanagedType.BStr)] string idMso);
}
