using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("l={l}, f={f}, cnt={cnt}")]
[CompilerGenerated]
internal sealed class K<A, B, C> : IEquatable<K<A, B, C>>
{
	private readonly A A;

	private readonly B A;

	private readonly C A;

	public A l => this.A;

	public B f => this.A;

	public C cnt => this.A;

	public K(A A, B B, C C)
	{
		this.A = A;
		this.A = B;
		this.A = C;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(491), new object[3] { this.A, this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = (-538928259 * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0)) * -1521134295;
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

	public bool Equals(K<A, B, C> val)
	{
		if (this != val)
		{
			while (true)
			{
				switch (3)
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
											switch (6)
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
								IL_00bd:
								if (num2 != 0)
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
								goto IL_0105;
								IL_0072:
								if (num)
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
									object obj = this.A;
									object obj2 = val.A;
									if (obj != null)
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
										if (obj2 != null)
										{
											num2 = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
											goto IL_00bd;
										}
									}
									num2 = obj == obj2;
									goto IL_00bd;
								}
								goto IL_0105;
								IL_0105:
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

	bool IEquatable<K<A, B, C>>.Equals(K<A, B, C> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as K<A, B, C>);
	}
}
