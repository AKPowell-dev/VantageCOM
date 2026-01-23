using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[Guid("86488FB4-9633-4C93-8057-FC1FA7A847AE")]
[CompilerGenerated]
public interface ChartGroup
{
	void _VtblGap1_4();

	[DispId(1610743812)]
	DownBars DownBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743812)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(1610743813)]
	DropLines DropLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743813)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_4();

	[DispId(1610743818)]
	bool HasDropLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743818)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743818)]
		[param: In]
		set;
	}

	[DispId(1610743820)]
	bool HasHiLoLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743820)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743820)]
		[param: In]
		set;
	}

	void _VtblGap3_4();

	[DispId(1610743826)]
	bool HasUpDownBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743826)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743826)]
		[param: In]
		set;
	}

	[DispId(1610743828)]
	HiLoLines HiLoLines
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743828)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap4_10();

	[DispId(1610743839)]
	UpBars UpBars
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743839)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
