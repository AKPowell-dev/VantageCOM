using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace A;

[CompilerGenerated]
internal sealed class TH
{
	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 5)]
	internal struct PH
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 16)]
	internal struct QH
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 32)]
	internal struct RH
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 36)]
	internal struct SH
	{
	}

	internal static readonly QH A/* Not supported: data(E0 EF FF FF B6 EF FF FF B5 EF FF FF BA EF FF FF) */;

	internal static readonly SH A/* Not supported: data(01 00 00 00 07 00 00 00 04 00 00 00 09 00 00 00 08 00 00 00 03 00 00 00 05 00 00 00 02 00 00 00 06 00 00 00) */;

	internal static readonly PH A/* Not supported: data(01 00 00 01 01) */;

	internal static readonly RH A/* Not supported: data(F4 EF FF FF 00 00 00 00 01 00 00 00 DD EF FF FF C8 EF FF FF 04 00 00 00 03 00 00 00 02 00 00 00) */;

	internal static readonly QH B/* Not supported: data(08 00 00 00 09 00 00 00 07 00 00 00 0A 00 00 00) */;

	internal static uint A(string A)
	{
		uint num = 2166136261u;
		if (A != null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			for (int i = 0; i < A.Length; i++)
			{
				num = (A[i] ^ num) * 16777619;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		return num;
	}
}
