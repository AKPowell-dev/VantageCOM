using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("Color={Color}, Count={Count}")]
[CompilerGenerated]
internal sealed class A<A, B> : IEquatable<A<A, B>>
{
	private readonly A m_A;

	private readonly B m_A;

	public A Color => this.A;

	public B Count => this.A;

	public A(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(1), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		return (355885746 * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0)) * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0);
	}

	public bool Equals(A<A, B> val)
	{
		if (this != val)
		{
			while (true)
			{
				bool num;
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
						if (val == null)
						{
							return false;
						}
						object obj = this.A;
						object obj2 = val.A;
						if (obj != null)
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
							if (obj2 != null)
							{
								num = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
								goto IL_0068;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						num = obj == obj2;
						goto IL_0068;
					}
					IL_0068:
					if (num)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
							{
								object obj = this.A;
								object obj2 = val.A;
								if (obj != null)
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
									if (obj2 != null)
									{
										return obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
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
								return obj == obj2;
							}
							}
						}
					}
					return false;
				}
			}
		}
		return true;
	}

	bool IEquatable<A<A, B>>.Equals(A<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as A<A, B>);
	}
}
