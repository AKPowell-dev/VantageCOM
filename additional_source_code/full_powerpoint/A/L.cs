using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("x={x}")]
[CompilerGenerated]
internal sealed class L<A> : IEquatable<L<A>>
{
	private readonly A A;

	public A x => this.A;

	public L(A A)
	{
		this.A = A;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(558), new object[1] { this.A });
	}

	public override int GetHashCode()
	{
		int num = 1505763070 * -1521134295;
		int num2;
		if (this.A != null)
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
			num2 = this.A.GetHashCode();
		}
		else
		{
			num2 = 0;
		}
		return num + num2;
	}

	public bool Equals(L<A> val)
	{
		if (this != val)
		{
			if (val != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						object obj = this.A;
						object obj2 = val.A;
						if (obj != null)
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
							if (obj2 != null)
							{
								return obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
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
						return obj == obj2;
					}
					}
				}
			}
			return false;
		}
		return true;
	}

	bool IEquatable<L<A>>.Equals(L<A> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as L<A>);
	}
}
