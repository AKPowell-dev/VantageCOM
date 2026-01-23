using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Excel;

[ComImport]
[CompilerGenerated]
[Guid("000C0398-0000-0000-C000-000000000046")]
[TypeIdentifier]
public interface TextFrame2 : _IMsoDispObj
{
	void _VtblGap1_23();

	[DispId(110)]
	MsoTriState WordWrap
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(110)]
		[param: In]
		set;
	}

	[DispId(111)]
	MsoAutoSize AutoSize
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		[param: In]
		set;
	}

	void _VtblGap2_2();

	[DispId(114)]
	TextRange2 TextRange
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(114)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
