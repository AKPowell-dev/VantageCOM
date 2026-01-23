using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[CompilerGenerated]
[DebuggerDisplay("PageNumber={PageNumber}, Count={Count}")]
internal sealed class G<A, B> : IEquatable<G<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A PageNumber => this.A;

	public B Count => this.A;

	public G(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, XC.A(270), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = 1885701116 * -1521134295;
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
		return (num + num2) * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0);
	}

	public bool Equals(G<A, B> val)
	{
		if (this != val)
		{
			while (true)
			{
				switch (4)
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
							switch (3)
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
										switch (7)
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
													switch (6)
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
