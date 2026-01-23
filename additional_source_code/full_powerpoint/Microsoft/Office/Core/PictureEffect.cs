using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core;

[ComImport]
[Guid("000C03D1-0000-0000-C000-000000000046")]
[TypeIdentifier]
[DefaultMember("Type")]
[CompilerGenerated]
public interface PictureEffect : _IMsoDispObj
{
	void _VtblGap1_2();

	[DispId(0)]
	MsoPictureEffectType Type
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(0)]
		get;
	}

	void _VtblGap2_2();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(2)]
	void Delete();

	[DispId(3)]
	EffectParameters EffectParameters
	{
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		[DispId(3)]
		[return: MarshalAs(UnmanagedType.Interface)]
		get;
	}
}
