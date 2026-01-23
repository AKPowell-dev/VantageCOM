using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[TypeIdentifier]
[Guid("000209C8-0000-0000-C000-000000000046")]
[CompilerGenerated]
public interface FillFormat
{
	void _VtblGap1_3();

	[DispId(100)]
	ColorFormat BackColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(100)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[DispId(101)]
	ColorFormat ForeColor
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(101)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap2_9();

	[DispId(111)]
	float Transparency
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(111)]
		[param: In]
		set;
	}

	void _VtblGap3_1();

	[DispId(113)]
	MsoTriState Visible
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(113)]
		[param: In]
		set;
	}
}
