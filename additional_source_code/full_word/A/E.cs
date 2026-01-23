using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace A;

[CompilerGenerated]
[DebuggerDisplay("Top={Top}, Right={Right}, Bottom={Bottom}, Left={Left}")]
internal sealed class E<A, B, C, D> : IEquatable<E<A, B, C, D>>
{
	private readonly A A;

	private readonly B A;

	private readonly C A;

	private readonly D A;

	public A Top => this.A;

	public B Right => this.A;

	public C Bottom => this.A;

	public D Left => this.A;

	public E(A A, B B, C C, D D)
	{
		this.A = A;
		this.A = B;
		this.A = C;
		this.A = D;
	}

	public override string ToString()
	{
		return string.Format(null, XC.A(104), this.A, this.A, this.A, this.A);
	}

	public override int GetHashCode()
	{
		int num = -1720826138 * -1521134295;
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
		return ((num3 + num4) * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0)) * -1521134295 + ((this.A != null) ? this.A.GetHashCode() : 0);
	}

	public bool Equals(E<A, B, C, D> val)
	{
		if (this != val)
		{
			if (val != null)
			{
				while (true)
				{
					bool num;
					bool num2;
					bool num3;
					switch (4)
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
									switch (7)
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
							num = obj == obj2;
							goto IL_0068;
						}
						IL_0106:
						if (num2 != 0)
						{
							while (true)
							{
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
											switch (3)
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
						}
						goto IL_0158;
						IL_0158:
						return false;
						IL_0068:
						if (num)
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
									num3 = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
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
							num3 = obj == obj2;
							goto IL_00bd;
						}
						goto IL_0158;
						IL_00bd:
						if (num3)
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
							object obj = this.A;
							object obj2 = val.A;
							if (obj != null)
							{
								if (obj2 != null)
								{
									num2 = obj.Equals(RuntimeHelpers.GetObjectValue(obj2));
									goto IL_0106;
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
							num2 = obj == obj2;
							goto IL_0106;
						}
						goto IL_0158;
					}
				}
			}
			return false;
		}
		return true;
	}

	bool IEquatable<E<A, B, C, D>>.Equals(E<A, B, C, D> val)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(val);
	}

	public override bool Equals(object obj)
	{
		return Equals(obj as E<A, B, C, D>);
	}
}
