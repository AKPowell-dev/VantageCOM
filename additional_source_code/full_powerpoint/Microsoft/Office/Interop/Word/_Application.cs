using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[DefaultMember("Name")]
[TypeIdentifier]
[Guid("00020970-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface _Application
{
	void _VtblGap1_3();

	[DispId(0)]
	string Name
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	void _VtblGap2_13();

	[DispId(14)]
	Languages Languages
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(14)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_8();

	[DispId(23)]
	bool Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(23)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(23)]
		[param: In]
		set;
	}

	void _VtblGap4_85();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1105)]
	void Quit([Optional][In][MarshalAs(UnmanagedType.Struct)] ref object SaveChanges, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object OriginalFormat, [Optional][In][MarshalAs(UnmanagedType.Struct)] ref object RouteDocument);
}
