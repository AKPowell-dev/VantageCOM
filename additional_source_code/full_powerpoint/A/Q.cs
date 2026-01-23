using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("sz={sz}, cnt={cnt}")]
[CompilerGenerated]
internal sealed class Q<A, B> : IEquatable<Q<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A sz => this.A;

	public B cnt => this.A;

	public Q(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(803), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = 575759740 * -1521134295;
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
		return (num + num2) * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0);
	}

	public bool Equals(Q<A, B> val)
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
							bool num;
							switch (7)
							{
							case 0:
								break;
							default:
								{
									object obj = this.A;
									object obj2 = val.A;
									if (obj != null)
									{
										if (obj2 != null)
										{
											num = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
											goto IL_0068;
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
									goto IL_0068;
								}
								IL_0068:
								if (num)
								{
									while (true)
									{
										switch (5)
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
													switch (2)
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
						}
					}
					return false;
				}
			}
		}
		return true;
	}

	bool IEquatable<Q<A, B>>.Equals(Q<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as Q<A, B>);
	}
}
