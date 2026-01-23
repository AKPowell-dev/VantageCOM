using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[TypeIdentifier]
[CompilerGenerated]
[Guid("914934B9-5A91-11CF-8700-00AA0060263B")]
public interface Tags : Collection
{
	void _VtblGap1_2();

	[DispId(11)]
	int Count
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(11)]
		get;
	}

	void _VtblGap2_2();

	[DispId(0)]
	string this[[In][MarshalAs(UnmanagedType.BStr)] string Name]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.BStr)]
		get;
	}

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2003)]
	void Add([In][MarshalAs(UnmanagedType.BStr)] string Name, [In][MarshalAs(UnmanagedType.BStr)] string Value);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2004)]
	void Delete([In][MarshalAs(UnmanagedType.BStr)] string Name);

	void _VtblGap3_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2007)]
	[return: MarshalAs(UnmanagedType.BStr)]
	string Name([In] int Index);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2008)]
	[return: MarshalAs(UnmanagedType.BStr)]
	string Value([In] int Index);
}
