using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint;

[ComImport]
[CompilerGenerated]
[Guid("914934DE-5A91-11CF-8700-00AA0060263B")]
[TypeIdentifier]
public interface Sequence : Collection
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
	Effect this[[In] int Index]
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}

	void _VtblGap3_9();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2012)]
	[return: MarshalAs(UnmanagedType.Interface)]
	Effect AddTriggerEffect([In][MarshalAs(UnmanagedType.Interface)] Shape pShape, [In] MsoAnimEffect effectId, [In] MsoAnimTriggerType trigger, [In][MarshalAs(UnmanagedType.Interface)] Shape pTriggerShape, [In][MarshalAs(UnmanagedType.BStr)] string bookmark = "", [In] MsoAnimateByLevel Level = MsoAnimateByLevel.msoAnimateLevelNone);
}
