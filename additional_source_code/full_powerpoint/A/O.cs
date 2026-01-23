using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("RelativeSize={RelativeSize}")]
[CompilerGenerated]
internal sealed class O<A> : IEquatable<O<A>>
{
	private readonly A A;

	public A RelativeSize => this.A;

	public O(A A)
	{
		this.A = A;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(695), new object[1] { this.A });
	}

	public override int GetHashCode()
	{
		int num = 467534989 * -1521134295;
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

	public bool Equals(O<A> val)
	{
		if (this != val)
		{
			if (val != null)
			{
				while (true)
				{
					switch (2)
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
								switch (7)
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
		return true;
	}

	bool IEquatable<O<A>>.Equals(O<A> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as O<A>);
	}
}
