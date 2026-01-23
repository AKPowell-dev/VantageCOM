using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[CompilerGenerated]
[Guid("000C0321-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface ThreeDFormat : _IMsoDispObj
{
	void _VtblGap1_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(12)]
	void ResetRotation();

	void _VtblGap2_21();

	[DispId(111)]
	MsoTriState Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		[param: In]
		set;
	}

	void _VtblGap3_8();

	[DispId(114)]
	MsoBevelType BevelTopType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(114)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(114)]
		[param: In]
		set;
	}

	void _VtblGap4_4();

	[DispId(117)]
	MsoBevelType BevelBottomType
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(117)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(117)]
		[param: In]
		set;
	}
}
