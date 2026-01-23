using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("punct={punct}, Group={Group}")]
[CompilerGenerated]
internal sealed class M<A, B> : IEquatable<M<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A punct => this.A;

	public B Group => this.A;

	public M(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(585), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = -1954825085 * -1521134295;
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
		int num3 = (num + num2) * -1521134295;
		int num4;
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
			num4 = this.A.GetHashCode();
		}
		else
		{
			num4 = 0;
		}
		return num3 + num4;
	}

	public bool Equals(M<A, B> val)
	{
		bool num;
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
					if (obj2 != null)
					{
						num = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
						goto IL_005c;
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
				num = obj == obj2;
				goto IL_005c;
			}
			return false;
		}
		return true;
		IL_005c:
		if (num)
		{
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
					switch (7)
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

	bool IEquatable<M<A, B>>.Equals(M<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as M<A, B>);
	}
}
