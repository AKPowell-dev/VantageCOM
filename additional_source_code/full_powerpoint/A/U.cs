using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[DebuggerDisplay("Top={Top}, Right={Right}, Bottom={Bottom}, Left={Left}")]
[CompilerGenerated]
internal sealed class U<A, B, C, D> : IEquatable<U<A, B, C, D>>
{
	private readonly A A;

	private readonly B A;

	private readonly C A;

	private readonly D A;

	public A Top => this.A;

	public B Right => this.A;

	public C Bottom => this.A;

	public D Left => this.A;

	public U(A A, B B, C C, D D)
	{
		this.A = A;
		this.A = B;
		this.A = C;
		this.A = D;
	}

	public override string ToString()
	{
		return string.Format(null, AH.A(1011), this.A, this.A, this.A, this.A);
	}

	public override int GetHashCode()
	{
		int num = (-1720826138 * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0)) * -1521134295;
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
		int num5 = (num3 + num4) * -1521134295;
		int num6;
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
			num6 = this.A.GetHashCode();
		}
		else
		{
			num6 = 0;
		}
		return num5 + num6;
	}

	public bool Equals(U<A, B, C, D> val)
	{
		if (this != val)
		{
			if (val != null)
			{
				while (true)
				{
					bool num;
					bool num2;
					switch (7)
					{
					case 0:
						break;
					default:
						{
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
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
									num = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
									goto IL_005a;
								}
							}
							num = obj == obj2;
							goto IL_005a;
						}
						IL_00ea:
						if (num2 != 0)
						{
							object obj = this.A;
							object obj2 = val.A;
							if (obj == null || obj2 == null)
							{
								return obj == obj2;
							}
							return obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
						}
						goto IL_0120;
						IL_0120:
						return false;
						IL_005a:
						if (num)
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
							object obj = this.A;
							object obj2 = val.A;
							if ((obj != null && obj2 != null) ? obj.Equals(RuntimeHelpers.GetObjectValue(obj2)) : (obj == obj2))
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
								obj = this.A;
								obj2 = val.A;
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
										num2 = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
										goto IL_00ea;
									}
								}
								num2 = obj == obj2;
								goto IL_00ea;
							}
						}
						goto IL_0120;
					}
				}
			}
			return false;
		}
		return true;
	}

	bool IEquatable<U<A, B, C, D>>.Equals(U<A, B, C, D> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as U<A, B, C, D>);
	}
}
