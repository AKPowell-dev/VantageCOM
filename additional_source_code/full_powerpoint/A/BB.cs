using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[CompilerGenerated]
[DebuggerDisplay("top={top}, Group={Group}")]
internal sealed class BB<A, B> : IEquatable<BB<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A top => this.A;

	public B Group => this.A;

	public BB(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(1449), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = (-1323111380 * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0)) * -1521134295;
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

	public bool Equals(BB<A, B> val)
	{
		if (this != val)
		{
			while (true)
			{
				switch (2)
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
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											break;
										}
										if (obj2 != null)
										{
											num = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
											goto IL_0068;
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
										switch (2)
										{
										case 0:
											break;
										default:
										{
											object obj = this.A;
											object obj2 = val.A;
											if (obj == null || obj2 == null)
											{
												return obj == obj2;
											}
											return obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
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

	bool IEquatable<BB<A, B>>.Equals(BB<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as BB<A, B>);
	}
}
