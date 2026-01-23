using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("f={f}, cnt={cnt}")]
[CompilerGenerated]
internal sealed class G<A, B> : IEquatable<G<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A f => this.A;

	public B cnt => this.A;

	public G(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(207), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = -1835808612 * -1521134295;
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
		int num3 = (num + num2) * -1521134295;
		int num4;
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
			num4 = this.A.GetHashCode();
		}
		else
		{
			num4 = 0;
		}
		return num3 + num4;
	}

	public bool Equals(G<A, B> val)
	{
		if (this != val)
		{
			while (true)
			{
				bool num;
				switch (5)
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
								switch (5)
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
							switch (2)
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
			}
		}
		return true;
	}

	bool IEquatable<G<A, B>>.Equals(G<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as G<A, B>);
	}
}
