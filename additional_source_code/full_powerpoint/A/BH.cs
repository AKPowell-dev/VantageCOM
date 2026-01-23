using System;
using System.IO;
using System.Reflection;

namespace A;

internal sealed class BH
{
	private static readonly Assembly m_A;

	static BH()
	{
		AppDomain.CurrentDomain.ResourceResolve += B;
		AppDomain.CurrentDomain.AssemblyResolve += A;
		Assembly executingAssembly = Assembly.GetExecutingAssembly();
		string assemblyString = A(executingAssembly);
		BH.m_A = Assembly.Load(assemblyString);
	}

	internal static void A()
	{
	}

	private static Assembly A(object A, ResolveEventArgs B)
	{
		Assembly executingAssembly = Assembly.GetExecutingAssembly();
		string text = BH.A(executingAssembly);
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
					byte[] rawAssembly = ZG.A(0, manifestResourceStream);
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
		if ((object)BH.m_A != null)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					string[] manifestResourceNames = BH.m_A.GetManifestResourceNames();
					foreach (string text in manifestResourceNames)
					{
						if (text == B.Name)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									return BH.m_A;
								}
							}
						}
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							return null;
						}
					}
				}
				}
			}
		}
		return BH.m_A;
	}
}
