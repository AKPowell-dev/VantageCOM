using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[Guid("91493462-5A91-11CF-8700-00AA0060263B")]
[CompilerGenerated]
[TypeIdentifier]
public interface Presentations : Collection
{
	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(-4)]
	[return: MarshalAs(UnmanagedType.CustomMarshaler, MarshalType = "System.Runtime.InteropServices.CustomMarshalers.EnumeratorToEnumVariantMarshaler, CustomMarshalers, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
	new IEnumerator GetEnumerator();

	void _VtblGap1_1();

	[DispId(11)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(11)]
		get;
	}

	void _VtblGap2_2();

	[DispId(0)]
	Presentation this[[In][MarshalAs(UnmanagedType.Struct)] object Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2003)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Presentation Add([In] MsoTriState WithWindow = MsoTriState.msoTrue);

	void _VtblGap3_1();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2005)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Presentation Open([In][MarshalAs(UnmanagedType.BStr)] string FileName, [In] MsoTriState ReadOnly = MsoTriState.msoFalse, [In] MsoTriState Untitled = MsoTriState.msoFalse, [In] MsoTriState WithWindow = MsoTriState.msoTrue);
}
