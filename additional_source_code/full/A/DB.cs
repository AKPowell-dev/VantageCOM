using System;

namespace A;

internal sealed class DB : Exception
{
	internal readonly Exception A;

	internal DB(Exception A)
	{
		this.A = A;
	}
}
