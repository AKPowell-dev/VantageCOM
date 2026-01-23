using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;

namespace A;

internal sealed class EH
{
	public sealed class DH
	{
		private static OpCode[] m_A;

		private static OpCode[] m_B;

		private int m_A;

		private byte[] m_A;

		private DynamicILInfo m_A;

		private Module m_A;

		private Type[] m_A;

		private Type[] m_B;

		static DH()
		{
			DH.m_A = new OpCode[256];
			DH.m_B = new OpCode[256];
			FieldInfo[] fields = typeof(OpCodes).GetFields(BindingFlags.Static | BindingFlags.Public);
			foreach (FieldInfo fieldInfo in fields)
			{
				OpCode opCode = (OpCode)fieldInfo.GetValue(null);
				ushort num = (ushort)opCode.Value;
				if (num < 256)
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
					DH.m_A[num] = opCode;
				}
				else if ((num & 0xFF00) == 65024)
				{
					DH.m_B[num & 0xFF] = opCode;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}

		public DH(MethodBase A, byte[] B, DynamicILInfo C)
		{
			this.m_A = C;
			this.m_A = B;
			this.m_A = 0;
			this.m_A = A.Module;
			object a;
			if (!(A is ConstructorInfo))
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
				a = A.GetGenericArguments();
			}
			else
			{
				a = null;
			}
			this.m_A = (Type[])a;
			object b;
			if ((object)A.DeclaringType != null)
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
				b = A.DeclaringType.GetGenericArguments();
			}
			else
			{
				b = null;
			}
			this.m_B = (Type[])b;
		}

		internal void A()
		{
			while (this.m_A < this.m_A.Length)
			{
				A();
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
				return;
			}
		}

		private object A()
		{
			int a = this.m_A;
			OpCode nop = OpCodes.Nop;
			int num = 0;
			byte b = A();
			if (b != 254)
			{
				nop = DH.m_A[b];
			}
			else
			{
				b = A();
				nop = DH.m_B[b];
			}
			switch (nop.OperandType)
			{
			case OperandType.InlineNone:
				return null;
			case OperandType.ShortInlineBrTarget:
				A(1);
				return null;
			case OperandType.InlineBrTarget:
				A(4);
				return null;
			case OperandType.ShortInlineI:
				A(1);
				return null;
			case OperandType.InlineI:
				A(4);
				return null;
			case OperandType.InlineI8:
				A(8);
				return null;
			case OperandType.ShortInlineR:
				A(4);
				return null;
			case OperandType.InlineR:
				A(8);
				return null;
			case OperandType.ShortInlineVar:
				A(1);
				return null;
			case OperandType.InlineVar:
				A(2);
				return null;
			case OperandType.InlineString:
				num = A();
				B(this.m_A.GetTokenFor(this.m_A.ResolveString(num)), a + nop.Size);
				return null;
			case OperandType.InlineSig:
				num = A();
				B(this.m_A.GetTokenFor(this.m_A.ResolveSignature(num)), a + nop.Size);
				return null;
			case OperandType.InlineMethod:
			{
				num = A();
				MethodBase methodBase2 = this.m_A.ResolveMethod(num, this.m_B, this.m_A);
				B(this.m_A.GetTokenFor(methodBase2.MethodHandle, methodBase2.DeclaringType.TypeHandle), a + nop.Size);
				return null;
			}
			case OperandType.InlineField:
			{
				num = A();
				FieldInfo fieldInfo2 = this.m_A.ResolveField(num, this.m_B, this.m_A);
				B(this.m_A.GetTokenFor(fieldInfo2.FieldHandle), a + nop.Size);
				return null;
			}
			case OperandType.InlineType:
			{
				num = A();
				Type type2 = this.m_A.ResolveType(num, this.m_B, this.m_A);
				B(this.m_A.GetTokenFor(type2.TypeHandle), a + nop.Size);
				return null;
			}
			case OperandType.InlineTok:
			{
				num = A();
				MemberInfo memberInfo = this.m_A.ResolveMember(num, this.m_B, this.m_A);
				if (memberInfo.MemberType != MemberTypes.TypeInfo)
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
					if (memberInfo.MemberType != MemberTypes.NestedType)
					{
						if (memberInfo.MemberType != MemberTypes.Method)
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
							if (memberInfo.MemberType != MemberTypes.Constructor)
							{
								if (memberInfo.MemberType == MemberTypes.Field)
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
									FieldInfo fieldInfo = memberInfo as FieldInfo;
									num = this.m_A.GetTokenFor(fieldInfo.FieldHandle);
								}
								goto IL_0373;
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
						MethodBase methodBase = memberInfo as MethodBase;
						num = this.m_A.GetTokenFor(methodBase.MethodHandle, methodBase.DeclaringType.TypeHandle);
						goto IL_0373;
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
				}
				Type type = memberInfo as Type;
				num = this.m_A.GetTokenFor(type.TypeHandle);
				goto IL_0373;
			}
			case OperandType.InlineSwitch:
			{
				int num2 = A();
				A(num2 * 4);
				return null;
			}
			default:
				{
					throw new BadImageFormatException("unexpected OperandType " + nop.OperandType);
				}
				IL_0373:
				B(num, a + nop.Size);
				return null;
			}
		}

		private void A(int A)
		{
			this.m_A += A;
		}

		private byte A()
		{
			return this.m_A[this.m_A++];
		}

		private int A()
		{
			int a = this.m_A;
			this.m_A += 4;
			return BitConverter.ToInt32(this.m_A, a);
		}

		private void B(int A, int B)
		{
			this.m_A[B++] = (byte)A;
			this.m_A[B++] = (byte)(A >> 8);
			this.m_A[B++] = (byte)(A >> 16);
			this.m_A[B++] = (byte)(A >> 24);
		}
	}

	internal static readonly byte[] A;

	internal static readonly Dictionary<int, int> A;

	private static readonly ModuleHandle m_A;

	static EH()
	{
		if (EH.A == null)
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
			string s = "TWFjYWJhY3VzLlBvd2VyUG9pbnQq";
			byte[] array = Convert.FromBase64String(s);
			s = Encoding.UTF8.GetString(array, 0, array.Length);
			Stream manifestResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(s);
			EH.A = ZG.A(0, manifestResourceStream);
			EH.A = new Dictionary<int, int>();
			BinaryReader binaryReader = new BinaryReader(new MemoryStream(EH.A, writable: false));
			try
			{
				int num = binaryReader.ReadInt32();
				for (int i = 0; i < num; i++)
				{
					int key = binaryReader.ReadInt32();
					int value = binaryReader.ReadInt32();
					EH.A[key] = value;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_00b0;
					}
					continue;
					end_IL_00b0:
					break;
				}
			}
			finally
			{
				if (binaryReader != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						((IDisposable)binaryReader).Dispose();
						break;
					}
				}
			}
		}
		if ((object)typeof(MulticastDelegate) == null)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			EH.m_A = Assembly.GetExecutingAssembly().GetModules()[0].ModuleHandle;
			return;
		}
	}

	internal static void A(int A, int B, int C)
	{
		Type typeFromHandle;
		MethodBase methodBase;
		try
		{
			typeFromHandle = Type.GetTypeFromHandle(EH.m_A.ResolveTypeHandle(A));
			object methodFromHandle = MethodBase.GetMethodFromHandle(EH.m_A.ResolveMethodHandle(B), EH.m_A.ResolveTypeHandle(C));
			methodBase = (MethodBase)methodFromHandle;
		}
		catch (Exception)
		{
			throw;
		}
		FieldInfo[] fields = typeFromHandle.GetFields(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.GetField);
		foreach (FieldInfo fieldInfo in fields)
		{
			try
			{
				DynamicMethod dynamicMethod = null;
				MethodBody methodBody = methodBase.GetMethodBody();
				Type[] parameterTypes = EH.A(methodBase);
				string name = methodBase.DeclaringType.FullName + "." + methodBase.Name + "_Encrypted$";
				object returnType;
				if (!(methodBase is ConstructorInfo))
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
					returnType = ((MethodInfo)methodBase).ReturnType;
				}
				else
				{
					returnType = null;
				}
				dynamicMethod = new DynamicMethod(name, (Type)returnType, parameterTypes, methodBase.DeclaringType, skipVisibility: true);
				EH.A.TryGetValue(A, out var value);
				DynamicILInfo dynamicILInfo = dynamicMethod.GetDynamicILInfo();
				EH.A(methodBody, dynamicILInfo);
				EH.A(ref value, methodBase, dynamicILInfo);
				EH.A(ref value, dynamicILInfo);
				Delegate value2 = dynamicMethod.CreateDelegate(typeFromHandle);
				fieldInfo.SetValue(null, value2);
			}
			catch (Exception)
			{
			}
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private static void A(ref int A, MethodBase B, DynamicILInfo C)
	{
		int maxStackSize = BitConverter.ToInt32(EH.A, A);
		A += 4;
		int num = BitConverter.ToInt32(EH.A, A);
		A += 4;
		byte[] array = new byte[num];
		Buffer.BlockCopy(EH.A, A, array, 0, num);
		DH dH = new DH(B, array, C);
		dH.A();
		C.SetCode(array, maxStackSize);
		A += num;
	}

	private static void A(MethodBody A, DynamicILInfo B)
	{
		SignatureHelper localVarSigHelper = SignatureHelper.GetLocalVarSigHelper();
		IEnumerator<LocalVariableInfo> enumerator = A.LocalVariables.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				LocalVariableInfo current = enumerator.Current;
				localVarSigHelper.AddArgument(current.LocalType, current.IsPinned);
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		B.SetLocalSignature(localVarSigHelper.GetSignature());
	}

	private static void A(ref int A, DynamicILInfo B)
	{
		int num = BitConverter.ToInt32(EH.A, A);
		A += 4;
		if (num == 0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return;
				}
			}
		}
		byte[] array = new byte[num];
		Buffer.BlockCopy(EH.A, A, array, 0, num);
		int num2 = 4;
		int num3 = (num - 4) / 24;
		for (int i = 0; i < num3; i++)
		{
			ExceptionHandlingClauseOptions exceptionHandlingClauseOptions = (ExceptionHandlingClauseOptions)BitConverter.ToInt32(array, num2);
			num2 += 20;
			switch (exceptionHandlingClauseOptions)
			{
			case ExceptionHandlingClauseOptions.Clause:
			{
				RuntimeTypeHandle type = EH.m_A.ResolveTypeHandle(BitConverter.ToInt32(array, num2));
				int tokenFor = B.GetTokenFor(type);
				EH.A(tokenFor, num2, array);
				break;
			}
			case ExceptionHandlingClauseOptions.Fault:
				throw new NotSupportedException("dynamic method does not support fault clause");
			}
			num2 += 4;
		}
		B.SetExceptions(array);
	}

	public static void A(int A, int B, byte[] C)
	{
		C[B++] = (byte)A;
		C[B++] = (byte)(A >> 8);
		C[B++] = (byte)(A >> 16);
		C[B++] = (byte)(A >> 24);
	}

	private static Type[] A(MethodBase A)
	{
		ParameterInfo[] parameters = A.GetParameters();
		int num = parameters.Length;
		if (!A.IsStatic)
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
			num++;
		}
		Type[] array = new Type[num];
		int num2 = 0;
		if (!A.IsStatic)
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
			if (A.DeclaringType.IsValueType)
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
				array[0] = A.DeclaringType.MakeByRefType();
			}
			else
			{
				array[0] = A.DeclaringType;
			}
			num2++;
		}
		int num3 = 0;
		while (num3 < parameters.Length)
		{
			array[num2] = parameters[num3].ParameterType;
			num3++;
			num2++;
		}
		return array;
	}
}
