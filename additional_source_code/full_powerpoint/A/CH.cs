using System;
using System.Collections;
using System.IO;
using System.Reflection;

namespace A;

internal sealed class CH
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
			text = new string(array);
			string[] array2 = text.Split(new string[1] { AH.A(170978) }, StringSplitOptions.RemoveEmptyEntries);
			if (array2 != null)
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
				if (array2.Length >= 0)
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
					for (int j = 0; j < array2.Length; j += 2)
					{
						if (array2[j + 1].StartsWith(AH.A(98013)))
						{
							try
							{
								Assembly executingAssembly = Assembly.GetExecutingAssembly();
								string path = Path.Combine(Path.GetDirectoryName(executingAssembly.Location), array2[j]);
								if (File.Exists(path))
								{
									continue;
								}
								string[] manifestResourceNames = executingAssembly.GetManifestResourceNames();
								foreach (string text2 in manifestResourceNames)
								{
									if (!(text2 == array2[j + 1]))
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
										break;
									}
									Stream manifestResourceStream = executingAssembly.GetManifestResourceStream(text2);
									byte[] array3 = ZG.A(0, manifestResourceStream);
									FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write);
									try
									{
										fileStream.Write(array3, 0, array3.Length);
									}
									finally
									{
										if (fileStream != null)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												((IDisposable)fileStream).Dispose();
												break;
											}
										}
									}
									manifestResourceStream.Close();
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_016d;
									}
									continue;
									end_IL_016d:
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
				}
			}
			AppDomain.CurrentDomain.AssemblyResolve += A;
			return;
		}
	}

	private static string A(byte[] A, string B, string C, string D)
	{
		B = Path.Combine(Path.GetTempPath(), B);
		string text = Path.Combine(B, C + D);
		if (!File.Exists(text))
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
		lock (CH.m_A)
		{
			Assembly assembly = null;
			string name = B.Name;
			string text = string.Empty;
			IEnumerator enumerator = CH.B.Keys.GetEnumerator();
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
							switch (2)
							{
							case 0:
								continue;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							assembly = CH.m_A[text2] as Assembly;
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
							text = CH.B[text2] as string;
							break;
						}
						break;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_00a4;
						}
						continue;
						end_IL_00a4:
						break;
					}
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			if (text.Length == 0)
			{
				while (true)
				{
					switch (4)
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
					switch (6)
					{
					case 0:
						continue;
					}
					Stream manifestResourceStream = executingAssembly.GetManifestResourceStream(text3);
					byte[] array = ZG.A(0, manifestResourceStream);
					byte[] array2 = null;
					try
					{
						text += AH.A(49917);
						string[] manifestResourceNames2 = executingAssembly.GetManifestResourceNames();
						foreach (string text4 in manifestResourceNames2)
						{
							if (!(text4 == text))
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
								break;
							}
							Stream manifestResourceStream2 = executingAssembly.GetManifestResourceStream(text4);
							array2 = ZG.A(0, manifestResourceStream2);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_019c;
							}
							continue;
							end_IL_019c:
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
								switch (4)
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
						string c = CH.A(name);
						string path = CH.A(array, text, c, AH.A(170983));
						if (array2 != null)
						{
							CH.A(array, text, c, AH.A(170992));
						}
						assembly = Assembly.LoadFile(path);
					}
					CH.m_A[name] = assembly;
					return assembly;
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				return null;
			}
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
