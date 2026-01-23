using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("Size={Size}")]
[CompilerGenerated]
internal sealed class D<A> : IEquatable<D<A>>
{
	private readonly A A;

	public A Size => this.A;

	public D(A A)
	{
		this.A = A;
	}

	public override string ToString()
	{
		return string.Format(null, XC.A(71), new object[1] { this.A });
	}

	public override int GetHashCode()
	{
		int num = -1947125907 * -1521134295;
		int num2;
		if (this.A != null)
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
			num2 = this.A.GetHashCode();
		}
		else
		{
			num2 = 0;
		}
		return num + num2;
	}

	public bool Equals(D<A> val)
	{
		if (this != val)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (val != null)
					{
						object obj = this.A;
						object obj2 = val.A;
						if (obj != null)
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
							if (obj2 != null)
							{
								return obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
							}
						}
						return obj == obj2;
					}
					return false;
				}
			}
		}
		return true;
	}

	bool IEquatable<D<A>>.Equals(D<A> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as D<A>);
	}
}
