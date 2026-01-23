using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("r={r}, cnt={cnt}")]
[CompilerGenerated]
internal sealed class IB<A, B> : IEquatable<IB<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A r => this.A;

	public B cnt => this.A;

	public IB(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(1836), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = 1784870733 * -1521134295;
		int num2;
		if (this.A != null)
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
			num2 = this.A.GetHashCode();
		}
		else
		{
			num2 = 0;
		}
		return (num + num2) * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0);
	}

	public bool Equals(IB<A, B> val)
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
							if (obj2 != null)
							{
								num = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
								goto IL_005e;
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
						num = obj == obj2;
						goto IL_005e;
					}
					IL_005e:
					if (num)
					{
						while (true)
						{
							switch (1)
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

	bool IEquatable<IB<A, B>>.Equals(IB<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as IB<A, B>);
	}
}
