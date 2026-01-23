using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace A;

internal sealed class AH
{
	internal static readonly byte[] A;

	static AH()
	{
		if (AH.A != null)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			string s = "TWFjYWJhY3VzLlBvd2VyUG9pbnQk";
			byte[] array = Convert.FromBase64String(s);
			s = Encoding.UTF8.GetString(array, 0, array.Length);
			Stream manifestResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(s);
			AH.A = ZG.A(0, manifestResourceStream);
			return;
		}
	}

	internal static string A(int A)
	{
		int num = 0;
		if ((AH.A[A] & 0x80) == 0)
		{
			num = AH.A[A];
			A++;
		}
		else if ((AH.A[A] & 0x40) == 0)
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
			num = (AH.A[A] & -129) << 8;
			num |= AH.A[A + 1];
			A += 2;
		}
		else
		{
			num = (AH.A[A] & -193) << 24;
			num |= AH.A[A + 1] << 16;
			num |= AH.A[A + 2] << 8;
			num |= AH.A[A + 3];
			A += 4;
		}
		if (num < 1)
		{
			return string.Empty;
		}
		string str = Encoding.Unicode.GetString(AH.A, A, num);
		return string.Intern(str);
	}
}
