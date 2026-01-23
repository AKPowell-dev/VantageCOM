using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace A;

[CompilerGenerated]
internal sealed class YG
{
	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 12)]
	internal struct PG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 16)]
	internal struct QG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 20)]
	internal struct RG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 28)]
	internal struct SG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 32)]
	internal struct TG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 36)]
	internal struct UG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 40)]
	internal struct VG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 44)]
	internal struct WG
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 48)]
	internal struct XG
	{
	}

	internal static readonly VG A/* Not supported: data(54 00 00 00 56 00 00 00 C9 EF FF FF 51 00 00 00 04 00 00 00 3F 00 00 00 40 00 00 00 41 00 00 00 42 00 00 00 43 00 00 00) */;

	internal static readonly QG A/* Not supported: data(51 00 00 00 41 00 00 00 42 00 00 00 43 00 00 00) */;

	internal static readonly XG A/* Not supported: data(00 00 00 00 00 00 20 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40) */;

	internal static readonly TG A/* Not supported: data(00 00 00 00 00 00 20 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40) */;

	internal static readonly RG A/* Not supported: data(76 00 00 00 7A 00 00 00 79 00 00 00 77 00 00 00 7B 00 00 00) */;

	internal static readonly TG B/* Not supported: data(00 00 00 00 00 00 10 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 40) */;

	internal static readonly TG C/* Not supported: data(53 00 00 00 54 00 00 00 55 00 00 00 56 00 00 00 77 00 00 00 7B 00 00 00 75 00 00 00 78 00 00 00) */;

	internal static readonly SG A/* Not supported: data(7B 00 00 00 78 00 00 00 76 00 00 00 7A 00 00 00 79 00 00 00 75 00 00 00 8C 00 00 00) */;

	internal static readonly PG A/* Not supported: data(53 00 00 00 55 00 00 00 52 00 00 00) */;

	internal static readonly WG A/* Not supported: data(77 00 00 00 75 00 00 00 53 00 00 00 54 00 00 00 55 00 00 00 56 00 00 00 E8 EF FF FF 05 00 00 00 FA EF FF FF 44 00 00 00 47 00 00 00) */;

	internal static readonly PG B/* Not supported: data(77 00 00 00 75 00 00 00 8C 00 00 00) */;

	internal static readonly UG A/* Not supported: data(05 00 00 00 FA EF FF FF 44 00 00 00 47 00 00 00 E8 EF FF FF 45 00 00 00 46 00 00 00 50 00 00 00 77 00 00 00) */;

	internal static readonly QG B/* Not supported: data(01 00 00 00 03 00 00 00 02 00 00 00 04 00 00 00) */;

	internal static uint A(string A)
	{
		uint num = 2166136261u;
		if (A != null)
		{
			while (true)
			{
				switch (6)
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
		}
		return num;
	}
}
