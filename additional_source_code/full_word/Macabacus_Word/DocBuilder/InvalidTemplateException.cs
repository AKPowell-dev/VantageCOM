using System;

namespace Macabacus_Word.DocBuilder;

public sealed class InvalidTemplateException : Exception
{
	public InvalidTemplateException()
	{
	}

	public InvalidTemplateException(string message)
		: base(message)
	{
	}

	public InvalidTemplateException(string message, Exception inner)
		: base(message, inner)
	{
	}
}
