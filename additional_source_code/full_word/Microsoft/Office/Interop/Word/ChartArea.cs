using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.Word;

[ComImport]
[CompilerGenerated]
[Guid("C75AD98A-74E9-49FE-8BF1-544839CC08A5")]
[TypeIdentifier]
public interface ChartArea
{
	void _VtblGap1_15();

	[DispId(127)]
	double Left
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(127)]
		[param: In]
		set;
	}

	[DispId(126)]
	double Top
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(126)]
		get;
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(126)]
		[param: In]
		set;
	}

	void _VtblGap2_4();

	[DispId(1610743831)]
	ChartFormat Format
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(1610743831)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
