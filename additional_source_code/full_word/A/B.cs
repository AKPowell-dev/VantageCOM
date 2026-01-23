using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("Format={Format}")]
[CompilerGenerated]
internal sealed class B<A> : IEquatable<B<A>>
{
	private readonly A A;

	public A Format => this.A;

	public B(A A)
	{
		this.A = A;
	}

	public override string ToString()
	{
		return string.Format(null, XC.A(1), new object[1] { this.A });
	}

	public override int GetHashCode()
	{
		int num = -243332744 * -1521134295;
		int num2;
		if (this.A != null)
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

	public bool Equals(B<A> val)
	{
		if (this != val)
		{
			if (val != null)
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
						object obj = this.A;
						object obj2 = val.A;
						if (obj != null)
						{
							if (obj2 != null)
							{
								return obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
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
						return obj == obj2;
					}
					}
				}
			}
			return false;
		}
		return true;
	}

	bool IEquatable<B<A>>.Equals(B<A> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as B<A>);
	}
}
