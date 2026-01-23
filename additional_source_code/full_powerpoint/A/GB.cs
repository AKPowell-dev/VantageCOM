using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("b={b}, cnt={cnt}")]
[CompilerGenerated]
internal sealed class GB<A, B> : IEquatable<GB<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A b => this.A;

	public B cnt => this.A;

	public GB(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(1726), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = 383370854 * -1521134295;
		int num2;
		if (this.A != null)
		{
			while (true)
			{
				switch (4)
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
				switch (2)
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

	public bool Equals(GB<A, B> val)
	{
		if (this != val)
		{
			while (true)
			{
				switch (5)
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
								if ((obj != null && obj2 != null) ? obj.Equals(RuntimeHelpers.GetObjectValue(obj2)) : (obj == obj2))
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											obj = this.A;
											obj2 = val.A;
											if (obj != null)
											{
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
								return false;
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

	bool IEquatable<GB<A, B>>.Equals(GB<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as GB<A, B>);
	}
}
