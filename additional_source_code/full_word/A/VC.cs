using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace A;

[CompilerGenerated]
internal sealed class VC
{
	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 5)]
	internal struct SC
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 6)]
	internal struct TC
	{
	}

	[StructLayout(LayoutKind.Explicit, Pack = 1, Size = 16)]
	internal struct UC
	{
	}

	internal static readonly int A/* Not supported: data(01 01 01 01) */;

	internal static readonly SC A/* Not supported: data(01 01 01 01 01) */;

	internal static readonly UC A/* Not supported: data(FF FF FF FF FC FF FF FF FD FF FF FF FE FF FF FF) */;

	internal static readonly TC A/* Not supported: data(01 01 01 01 01 01) */;

	internal static readonly int B/* Not supported: data(01 01 01 00) */;

	internal static uint A(string A)
	{
		uint num = 2166136261u;
		if (A != null)
		{
			while (true)
			{
				switch (7)
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
				switch (1)
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
