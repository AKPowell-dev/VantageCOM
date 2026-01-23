using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;

namespace A;

internal sealed class UH
{
	private static readonly object m_A;

	private static readonly int m_A;

	private static readonly int B;

	private static readonly MemoryStream m_A;

	private static readonly MemoryStream B;

	static UH()
	{
		UH.m_A = null;
		B = null;
		UH.m_A = int.MaxValue;
		UH.B = 3279872;
		UH.m_A = new MemoryStream(0);
		B = new MemoryStream(0);
		UH.m_A = new object();
	}

	private static string A(Assembly A)
	{
		string text = A.FullName;
		int num = text.IndexOf(',');
		if (num >= 0)
		{
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
				num = fullName.IndexOf("publickeytoken=");
			}
			if (num < 0)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						return null;
					}
				}
			}
			num += 15;
			if (fullName[num] != 'n')
			{
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
					switch (5)
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
		lock (UH.m_A)
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
					int num2 = 0;
					while (true)
					{
						if (num2 < array3.Length)
						{
							if (array3[num2] != 0)
							{
								flag = false;
								break;
							}
							num2++;
							continue;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						break;
					}
					if (flag)
					{
						array2 = UH.A(Assembly.GetExecutingAssembly());
					}
					dESCryptoServiceProvider.Key = array2;
					if (UH.m_A == null)
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
						if (UH.m_A == int.MaxValue)
						{
							UH.m_A.Capacity = (int)B.Length;
						}
						else
						{
							UH.m_A.Capacity = UH.m_A;
						}
					}
					UH.m_A.Position = 0L;
					ICryptoTransform cryptoTransform = dESCryptoServiceProvider.CreateDecryptor();
					int inputBlockSize = cryptoTransform.InputBlockSize;
					_ = cryptoTransform.OutputBlockSize;
					byte[] array4 = new byte[cryptoTransform.OutputBlockSize];
					byte[] array5 = new byte[cryptoTransform.InputBlockSize];
					int j;
					for (j = (int)B.Position; j + inputBlockSize < B.Length; j += inputBlockSize)
					{
						B.Read(array5, 0, inputBlockSize);
						int count = cryptoTransform.TransformBlock(array5, 0, inputBlockSize, array4, 0);
						UH.m_A.Write(array4, 0, count);
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
					B.Read(array5, 0, (int)(B.Length - j));
					byte[] array6 = cryptoTransform.TransformFinalBlock(array5, 0, (int)(B.Length - j));
					UH.m_A.Write(array6, 0, array6.Length);
					stream = UH.m_A;
					stream.Position = 0L;
					memoryStream = UH.m_A;
				}
				if ((num & 8) != 0)
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
					try
					{
						if (UH.B == null)
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
							if (UH.B == int.MinValue)
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
								UH.B.Capacity = (int)stream.Length * 2;
							}
							else
							{
								UH.B.Capacity = UH.B;
							}
						}
						UH.B.Position = 0L;
						DeflateStream deflateStream = new DeflateStream(stream, CompressionMode.Decompress);
						int num3 = 1000;
						byte[] buffer = new byte[num3];
						int num4;
						do
						{
							num4 = deflateStream.Read(buffer, 0, num3);
							if (num4 <= 0)
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
								break;
							}
							UH.B.Write(buffer, 0, num4);
						}
						while (num4 >= num3);
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							memoryStream = UH.B;
							break;
						}
					}
					catch (Exception)
					{
					}
				}
				if (memoryStream != null)
				{
					return memoryStream.ToArray();
				}
				byte[] array7 = new byte[B.Length - B.Position];
				B.Read(array7, 0, array7.Length);
				return array7;
			}
		}
	}
}
