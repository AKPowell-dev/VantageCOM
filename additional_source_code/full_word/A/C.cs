using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[CompilerGenerated]
[DebuggerDisplay("Font={Font}")]
internal sealed class C<A> : IEquatable<C<A>>
{
	private readonly A A;

	public A Font => this.A;

	public C(A A)
	{
		this.A = A;
	}

	public override string ToString()
	{
		return string.Format(null, XC.A(38), new object[1] { this.A });
	}

	public override int GetHashCode()
	{
		int num = -447674137 * -1521134295;
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

	public bool Equals(C<A> val)
	{
		if (this != val)
		{
			if (val != null)
			{
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
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
			return false;
		}
		return true;
	}

	bool IEquatable<C<A>>.Equals(C<A> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as C<A>);
	}
}
