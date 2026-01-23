using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("Before={Before}, After={After}, Within={Within}")]
[CompilerGenerated]
internal sealed class R<A, B, C> : IEquatable<R<A, B, C>>
{
	private readonly A A;

	private readonly B A;

	private readonly C A;

	public A Before => this.A;

	public B After => this.A;

	public C Within => this.A;

	public R(A A, B B, C C)
	{
		this.A = A;
		this.A = B;
		this.A = C;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(854), new object[3] { this.A, this.A, this.A });
	}

	public override int GetHashCode()
	{
		int num = ((2070964630 * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0)) * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0)) * -1521134295;
		int num2;
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

	public bool Equals(R<A, B, C> val)
	{
		if (this != val)
		{
			while (true)
			{
				switch (7)
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
							switch (4)
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
											continue;
										}
										break;
									}
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

	bool IEquatable<R<A, B, C>>.Equals(R<A, B, C> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as R<A, B, C>);
	}
}
