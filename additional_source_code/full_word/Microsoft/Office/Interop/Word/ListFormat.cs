using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("000209C0-0000-0000-C000-000000000046")]
public interface ListFormat
{
	[DispId(68)]
	int ListLevelNumber
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(68)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(68)]
		[param: In]
		set;
	}

	void _VtblGap1_5();

	[DispId(74)]
	WdListType ListType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(74)]
		get;
	}

	void _VtblGap2_17();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(215)]
	void ApplyListTemplate([In][MarshalAs(UnmanagedType.Interface)] ListTemplate ListTemplate, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ContinuePreviousList, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object ApplyTo, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object DefaultListBehavior);
}
