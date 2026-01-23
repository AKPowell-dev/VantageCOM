using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("left={left}, first={first}, Group={Group}")]
[CompilerGenerated]
internal sealed class J<A, B, C> : IEquatable<J<A, B, C>>
{
	private readonly A A;

	private readonly B A;

	private readonly C A;

	public A left => this.A;

	public B first => this.A;

	public C Group => this.A;

	public J(A A, B B, C C)
	{
		this.A = A;
		this.A = B;
		this.A = C;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(406), new object[3] { this.A, this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = 2086187842 * -1521134295;
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
		int num3 = (num + num2) * -1521134295;
		int num4;
		if (this.A != null)
		{
			while (true)
			{
				switch (3)
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
		int num5 = (num3 + num4) * -1521134295;
		int num6;
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
			num6 = this.A.GetHashCode();
		}
		else
		{
			num6 = 0;
		}
		return num5 + num6;
	}

	public bool Equals(J<A, B, C> val)
	{
		if (this != val)
		{
			while (true)
			{
				switch (1)
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
							bool num2;
							switch (1)
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
											goto IL_0072;
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
									num = obj == obj2;
									goto IL_0072;
								}
								IL_00bd:
								if (num2 != 0)
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
												while (true)
												{
													switch (1)
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
								goto IL_0111;
								IL_0072:
								if (num)
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										break;
									}
									object obj = this.A;
									object obj2 = val.A;
									if (obj != null)
									{
										if (obj2 != null)
										{
											num2 = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
											goto IL_00bd;
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
									num2 = obj == obj2;
									goto IL_00bd;
								}
								goto IL_0111;
								IL_0111:
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

	bool IEquatable<J<A, B, C>>.Equals(J<A, B, C> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as J<A, B, C>);
	}
}
