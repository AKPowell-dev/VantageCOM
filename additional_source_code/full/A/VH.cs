using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace A;

internal sealed class VH
{
	internal static readonly byte[] A;

	static VH()
	{
		if (VH.A != null)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			string s = "TWFjYWJhY3VzLkV4Y2VsJA==";
			byte[] array = Convert.FromBase64String(s);
			s = Encoding.UTF8.GetString(array, 0, array.Length);
			Stream manifestResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(s);
			VH.A = UH.A(0, manifestResourceStream);
			return;
		}
	}

	internal static string A(int A)
	{
		int num = 0;
		if ((VH.A[A] & 0x80) == 0)
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
			num = VH.A[A];
			A++;
		}
		else if ((VH.A[A] & 0x40) == 0)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			num = (VH.A[A] & -129) << 8;
			num |= VH.A[A + 1];
			A += 2;
		}
		else
		{
			num = (VH.A[A] & -193) << 24;
			num |= VH.A[A + 1] << 16;
			num |= VH.A[A + 2] << 8;
			num |= VH.A[A + 3];
			A += 4;
		}
		if (num < 1)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return string.Empty;
				}
			}
		}
		string str = Encoding.Unicode.GetString(VH.A, A, num);
		return string.Intern(str);
	}
}
