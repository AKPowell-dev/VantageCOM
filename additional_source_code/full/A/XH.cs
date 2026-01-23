using System;
using System.Collections;
using System.IO;
using System.Reflection;

namespace A;

internal sealed class XH
{
	private static readonly Hashtable m_A = new Hashtable();

	private static readonly Hashtable B = new Hashtable();

	internal static void A()
	{
		string text = "ﾲﾞﾜﾞﾝﾞﾜﾊﾌ\uffd1ﾼﾐﾒﾒﾐﾑￓ\uffdfﾩﾚﾍﾌﾖﾐﾑￂￆ\uffd1ￇ\uffd1ￍ\uffd1ￏￓ\uffdfﾼﾊﾓﾋﾊﾍﾚￂﾑﾚﾊﾋﾍﾞﾓￓ\uffdfﾯﾊﾝﾓﾖﾜﾴﾚﾆﾫﾐﾔﾚﾑￂￎￋￏﾝￌￋￇￊￍﾚￊﾜﾞￊﾚﾙￜￜﾆﾩﾌￔﾪﾬￆﾛﾮﾫﾞﾆﾔﾵﾒﾝﾓﾇﾨﾮﾔﾘￂￂￜￜﾲﾞﾜﾱﾊﾒﾚﾍﾖﾜﾪﾏﾻﾐﾈﾑￓ\uffdfﾩﾚﾍﾌﾖﾐﾑￂￎ\uffd1ￏ\uffd1ￎ\uffd1ￏￓ\uffdfﾼﾊﾓﾋﾊﾍﾚￂﾑﾚﾊﾋﾍﾞﾓￓ\uffdfﾯﾊﾝﾓﾖﾜﾴﾚﾆﾫﾐﾔﾚﾑￂￎￋￏﾝￌￋￇￊￍﾚￊﾜﾞￊﾚﾙￜￜﾧￇﾋ\uffc9ﾆﾏﾋￔﾪﾔﾌﾆￔﾆﾓﾧﾧﾦￊￊﾮﾮￂￂￜￜ";
		char[] array = text.ToCharArray();
		for (int i = 0; i < array.Length; i++)
		{
			array[i] = (char)(~(uint)array[i]);
		}
		text = new string(array);
		string[] array2 = text.Split(new string[1] { VH.A(212263) }, StringSplitOptions.RemoveEmptyEntries);
		if (array2 != null && array2.Length >= 0)
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
			for (int j = 0; j < array2.Length; j += 2)
			{
				if (array2[j + 1].StartsWith(VH.A(212268)))
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					try
					{
						Assembly executingAssembly = Assembly.GetExecutingAssembly();
						string path = Path.Combine(Path.GetDirectoryName(executingAssembly.Location), array2[j]);
						if (File.Exists(path))
						{
							continue;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							string[] manifestResourceNames = executingAssembly.GetManifestResourceNames();
							foreach (string text2 in manifestResourceNames)
							{
								if (text2 == array2[j + 1])
								{
									Stream manifestResourceStream = executingAssembly.GetManifestResourceStream(text2);
									byte[] array3 = UH.A(0, manifestResourceStream);
									using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
									{
										fileStream.Write(array3, 0, array3.Length);
									}
									manifestResourceStream.Close();
								}
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0157;
								}
								continue;
								end_IL_0157:
								break;
							}
							break;
						}
					}
					catch
					{
					}
				}
				else
				{
					B[array2[j]] = array2[j + 1];
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		AppDomain.CurrentDomain.AssemblyResolve += A;
	}

	private static string A(byte[] A, string B, string C, string D)
	{
		B = Path.Combine(Path.GetTempPath(), B);
		string text = Path.Combine(B, C + D);
		if (!File.Exists(text))
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
			Directory.CreateDirectory(B);
			FileStream fileStream = new FileStream(text, FileMode.Create, FileAccess.Write);
			fileStream.Write(A, 0, A.Length);
			fileStream.Close();
		}
		return text;
	}

	private static Assembly A(object A, ResolveEventArgs B)
	{
		lock (XH.m_A)
		{
			Assembly assembly = null;
			string name = B.Name;
			string text = string.Empty;
			IEnumerator enumerator = XH.B.Keys.GetEnumerator();
			try
			{
				while (true)
				{
					if (enumerator.MoveNext())
					{
						string text2 = (string)enumerator.Current;
						if (!text2.StartsWith(name))
						{
							continue;
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
							assembly = XH.m_A[text2] as Assembly;
							if ((object)assembly != null)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										return assembly;
									}
								}
							}
							text = XH.B[text2] as string;
							break;
						}
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_00a6;
						}
						continue;
						end_IL_00a6:
						break;
					}
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable disposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						disposable.Dispose();
						break;
					}
				}
			}
			if (text.Length == 0)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return null;
					}
				}
			}
			Assembly executingAssembly = Assembly.GetExecutingAssembly();
			string[] manifestResourceNames = executingAssembly.GetManifestResourceNames();
			foreach (string text3 in manifestResourceNames)
			{
				if (!(text3 == text))
				{
					continue;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					Stream manifestResourceStream = executingAssembly.GetManifestResourceStream(text3);
					byte[] array = UH.A(0, manifestResourceStream);
					byte[] array2 = null;
					try
					{
						text += VH.A(49303);
						string[] manifestResourceNames2 = executingAssembly.GetManifestResourceNames();
						foreach (string text4 in manifestResourceNames2)
						{
							if (!(text4 == text))
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
								break;
							}
							Stream manifestResourceStream2 = executingAssembly.GetManifestResourceStream(text4);
							array2 = UH.A(0, manifestResourceStream2);
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_01a8;
							}
							continue;
							end_IL_01a8:
							break;
						}
					}
					catch (Exception)
					{
					}
					bool flag = false;
					try
					{
						if (array2 == null)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								assembly = Assembly.Load(array);
								break;
							}
						}
						else
						{
							try
							{
								assembly = Assembly.Load(array, array2);
							}
							catch (Exception)
							{
								assembly = Assembly.Load(array);
							}
						}
					}
					catch (FileLoadException)
					{
						flag = true;
					}
					catch (BadImageFormatException)
					{
						flag = true;
					}
					if (flag)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						string c = XH.A(name);
						string path = XH.A(array, text, c, VH.A(212271));
						if (array2 != null)
						{
							XH.A(array, text, c, VH.A(212280));
						}
						assembly = Assembly.LoadFile(path);
					}
					XH.m_A[name] = assembly;
					return assembly;
				}
			}
			return null;
		}
	}

	private static string A(string A)
	{
		int num = A.IndexOf(',');
		if (num >= 0)
		{
			while (true)
			{
				switch (3)
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
			A = A.Substring(0, num);
		}
		return A;
	}
}
