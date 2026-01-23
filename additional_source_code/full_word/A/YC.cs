using System;
using System.IO;
using System.Reflection;

namespace A;

internal sealed class YC
{
	private static readonly Assembly m_A;

	static YC()
	{
		AppDomain.CurrentDomain.ResourceResolve += B;
		AppDomain.CurrentDomain.AssemblyResolve += A;
		Assembly executingAssembly = Assembly.GetExecutingAssembly();
		string assemblyString = A(executingAssembly);
		YC.m_A = Assembly.Load(assemblyString);
	}

	internal static void A()
	{
	}

	private static Assembly A(object A, ResolveEventArgs B)
	{
		Assembly executingAssembly = Assembly.GetExecutingAssembly();
		string text = YC.A(executingAssembly);
		if (B.Name.StartsWith(text))
		{
			Stream manifestResourceStream = executingAssembly.GetManifestResourceStream(text);
			byte[] rawAssembly = WC.A(0, manifestResourceStream);
			return Assembly.Load(rawAssembly);
		}
		return null;
	}

	private static string A(Assembly A)
	{
		string text = A.FullName;
		int num = text.IndexOf(',');
		if (num >= 0)
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
			text = text.Substring(0, num);
		}
		return text + '&';
	}

	private static Assembly B(object A, ResolveEventArgs B)
	{
		if ((object)YC.m_A != null)
		{
			string[] manifestResourceNames = YC.m_A.GetManifestResourceNames();
			foreach (string text in manifestResourceNames)
			{
				if (!(text == B.Name))
				{
					continue;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return YC.m_A;
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				return null;
			}
		}
		return YC.m_A;
	}
}
