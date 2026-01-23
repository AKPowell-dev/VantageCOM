using System;

namespace Macabacus_Word;

public sealed class MultipleBulletsInSmartArtException : Exception
{
	public MultipleBulletsInSmartArtException()
	{
	}

	public MultipleBulletsInSmartArtException(string message)
		: base(message)
	{
	}

	public MultipleBulletsInSmartArtException(string message, Exception inner)
		: base(message, inner)
	{
	}
}
