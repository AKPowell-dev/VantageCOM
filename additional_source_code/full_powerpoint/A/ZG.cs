using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;

namespace A;

internal sealed class ZG
{
	private static readonly object m_A;

	private static readonly int m_A;

	private static readonly int B;

	private static readonly MemoryStream m_A;

	private static readonly MemoryStream B;

	static ZG()
	{
		ZG.m_A = null;
		B = null;
		ZG.m_A = int.MaxValue;
		ZG.B = 3279872;
		ZG.m_A = new MemoryStream(0);
		B = new MemoryStream(0);
		ZG.m_A = new object();
	}

	private static string A(Assembly A)
	{
		string text = A.FullName;
		int num = text.IndexOf(',');
		if (num >= 0)
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
			text = text.Substring(0, num);
		}
		return text;
	}

	private static byte[] A(Assembly A)
	{
		try
		{
			string fullName = A.FullName;
			int num = fullName.IndexOf("PublicKeyToken=");
			if (num < 0)
			{
				num = fullName.IndexOf("publickeytoken=");
			}
			if (num < 0)
			{
				return null;
			}
			num += 15;
			if (fullName[num] != 'n')
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
				if (fullName[num] != 'N')
				{
					string s = fullName.Substring(num, 16);
					long value = long.Parse(s, NumberStyles.HexNumber);
					byte[] bytes = BitConverter.GetBytes(value);
					Array.Reverse(bytes);
					return bytes;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			return null;
		}
		catch
		{
		}
		return null;
	}

	internal static byte[] A(sbyte A, Stream B)
	{
		lock (ZG.m_A)
		{
			Stream stream = B;
			MemoryStream memoryStream = null;
			ushort num = (ushort)B.ReadByte();
			num = (ushort)(~num);
			for (int i = 1; i < 3; i++)
			{
				B.ReadByte();
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
				if ((num & 2) != 0)
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
					DESCryptoServiceProvider dESCryptoServiceProvider = new DESCryptoServiceProvider();
					byte[] array = new byte[8];
					B.Read(array, 0, 8);
					dESCryptoServiceProvider.IV = array;
					byte[] array2 = new byte[8];
					B.Read(array2, 0, 8);
					bool flag = true;
					byte[] array3 = array2;
					for (int j = 0; j < array3.Length; j++)
					{
						if (array3[j] == 0)
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
						flag = false;
						break;
					}
					if (flag)
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
						array2 = ZG.A(Assembly.GetExecutingAssembly());
					}
					dESCryptoServiceProvider.Key = array2;
					if (ZG.m_A == null)
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
						if (ZG.m_A == int.MaxValue)
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
							ZG.m_A.Capacity = (int)B.Length;
						}
						else
						{
							ZG.m_A.Capacity = ZG.m_A;
						}
					}
					ZG.m_A.Position = 0L;
					ICryptoTransform cryptoTransform = dESCryptoServiceProvider.CreateDecryptor();
					int inputBlockSize = cryptoTransform.InputBlockSize;
					_ = cryptoTransform.OutputBlockSize;
					byte[] array4 = new byte[cryptoTransform.OutputBlockSize];
					byte[] array5 = new byte[cryptoTransform.InputBlockSize];
					int k;
					for (k = (int)B.Position; k + inputBlockSize < B.Length; k += inputBlockSize)
					{
						B.Read(array5, 0, inputBlockSize);
						int count = cryptoTransform.TransformBlock(array5, 0, inputBlockSize, array4, 0);
						ZG.m_A.Write(array4, 0, count);
					}
					B.Read(array5, 0, (int)(B.Length - k));
					byte[] array6 = cryptoTransform.TransformFinalBlock(array5, 0, (int)(B.Length - k));
					ZG.m_A.Write(array6, 0, array6.Length);
					stream = ZG.m_A;
					stream.Position = 0L;
					memoryStream = ZG.m_A;
				}
				if ((num & 8) != 0)
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
					try
					{
						if (ZG.B == null)
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
							if (ZG.B == int.MinValue)
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
								ZG.B.Capacity = (int)stream.Length * 2;
							}
							else
							{
								ZG.B.Capacity = ZG.B;
							}
						}
						ZG.B.Position = 0L;
						DeflateStream deflateStream = new DeflateStream(stream, CompressionMode.Decompress);
						int num2 = 1000;
						byte[] buffer = new byte[num2];
						int num3;
						do
						{
							num3 = deflateStream.Read(buffer, 0, num2);
							if (num3 <= 0)
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
							ZG.B.Write(buffer, 0, num3);
						}
						while (num3 >= num2);
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							memoryStream = ZG.B;
							break;
						}
					}
					catch (Exception)
					{
					}
				}
				if (memoryStream != null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return memoryStream.ToArray();
						}
					}
				}
				byte[] array7 = new byte[B.Length - B.Position];
				B.Read(array7, 0, array7.Length);
				return array7;
			}
		}
	}
}
