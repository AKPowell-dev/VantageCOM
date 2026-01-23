using System;
using System.IO;
using System.Reflection;

namespace A;

internal sealed class WH
{
	private static readonly Assembly m_A;

	static WH()
	{
		AppDomain.CurrentDomain.ResourceResolve += B;
		AppDomain.CurrentDomain.AssemblyResolve += A;
		Assembly executingAssembly = Assembly.GetExecutingAssembly();
		string assemblyString = A(executingAssembly);
		WH.m_A = Assembly.Load(assemblyString);
	}

	internal static void A()
	{
	}

	private static Assembly A(object A, ResolveEventArgs B)
	{
		Assembly executingAssembly = Assembly.GetExecutingAssembly();
		string text = WH.A(executingAssembly);
		if (B.Name.StartsWith(text))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Stream manifestResourceStream = executingAssembly.GetManifestResourceStream(text);
					byte[] rawAssembly = UH.A(0, manifestResourceStream);
					return Assembly.Load(rawAssembly);
				}
				}
			}
		}
		return null;
	}

	private static string A(Assembly A)
	{
		string text = A.FullName;
		int num = text.IndexOf(',');
		if (num >= 0)
		{
			text = text.Substring(0, num);
		}
		return text + '&';
	}

	private static Assembly B(object A, ResolveEventArgs B)
	{
		if ((object)WH.m_A != null)
		{
			string[] manifestResourceNames = WH.m_A.GetManifestResourceNames();
			foreach (string text in manifestResourceNames)
			{
				if (!(text == B.Name))
				{
					continue;
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
					return WH.m_A;
				}
			}
			return null;
		}
		return WH.m_A;
	}
}
