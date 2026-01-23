using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace A;

internal sealed class XC
{
	internal static readonly byte[] A;

	static XC()
	{
		if (XC.A != null)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			string s = "TWFjYWJhY3VzLldvcmQk";
			byte[] array = Convert.FromBase64String(s);
			s = Encoding.UTF8.GetString(array, 0, array.Length);
			Stream manifestResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(s);
			XC.A = WC.A(0, manifestResourceStream);
			return;
		}
	}

	internal static string A(int A)
	{
		int num = 0;
		if ((XC.A[A] & 0x80) == 0)
		{
			while (true)
			{
				switch (1)
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
			num = XC.A[A];
			A++;
		}
		else if ((XC.A[A] & 0x40) == 0)
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
			num = (XC.A[A] & -129) << 8;
			num |= XC.A[A + 1];
			A += 2;
		}
		else
		{
			num = (XC.A[A] & -193) << 24;
			num |= XC.A[A + 1] << 16;
			num |= XC.A[A + 2] << 8;
			num |= XC.A[A + 3];
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
		string str = Encoding.Unicode.GetString(XC.A, A, num);
		return string.Intern(str);
	}
}
