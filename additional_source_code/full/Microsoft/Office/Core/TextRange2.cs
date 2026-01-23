using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[TypeIdentifier]
[DefaultMember("Text")]
[CompilerGenerated]
[Guid("000C0397-0000-0000-C000-000000000046")]
public interface TextRange2 : _IMsoDispObj, IEnumerable
{
	void _VtblGap1_2();

	[DispId(0)]
	string Text
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[param: In]
		[param: MarshalAs(UnmanagedType.BStr)]
		set;
	}

	void _VtblGap2_11();

	[DispId(11)]
	Font2 Font
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(11)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_27();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(38)]
	[return: MarshalAs(UnmanagedType.Interface)]
	TextRange2 InsertChartField([In] MsoChartFieldType ChartFieldType, [In][MarshalAs(UnmanagedType.BStr)] string Formula = "", [In] int Position = -1);
}
