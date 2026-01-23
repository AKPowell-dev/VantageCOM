using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("LeftIndent={LeftIndent}, FirstLineIndent={FirstLineIndent}")]
[CompilerGenerated]
internal sealed class H<A, B> : IEquatable<H<A, B>>
{
	private readonly A A;

	private readonly B A;

	public A LeftIndent => this.A;

	public B FirstLineIndent => this.A;

	public H(A A, B B)
	{
		this.A = A;
		this.A = B;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(256), new object[2] { this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = 1707614804 * -1521134295;
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
				switch (4)
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

	public bool Equals(H<A, B> val)
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
							bool num;
							switch (3)
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
											num = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
											goto IL_0072;
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
									goto IL_0072;
								}
								IL_0072:
								if (num)
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

	bool IEquatable<H<A, B>>.Equals(H<A, B> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as H<A, B>);
	}
}
