using System;
using System.Collections.Generic;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane.UI;
using MacabacusMacros.Libraries.Versioning;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Library2.Versioning;

public sealed class Tagging
{
	internal static readonly string A = AH.A(58984);

	internal static void A(Shape A, ContentItem B, string C, string D)
	{
		if (B.Id.Length <= 0)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				A.Tags.Add(Tagging.A, Content.GenerateXml(B, C, D, true));
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	internal static void A(Slide A, ContentItem B, string C, string D, bool E)
	{
		if (B.Id.Length > 0)
		{
			try
			{
				A.Tags.Add(Tagging.A, Content.GenerateXml(B, C, D, E));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	internal static ContentInfo? A(Shape A)
	{
		return Tagging.A(A.Tags);
	}

	internal static ContentInfo? A(Slide A)
	{
		return Tagging.A(A.Tags);
	}

	private static ContentInfo? A(Tags A)
	{
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		string text = Tagging.A(A);
		if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			XmlDocument xmlDocument = new XmlDocument();
			try
			{
				xmlDocument.LoadXml(text);
				return Content.GetContentInfoFromXml(xmlDocument);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				xmlDocument = null;
			}
		}
		return null;
	}

	internal static string A(Tags A)
	{
		return A[Tagging.A];
	}

	internal static bool A(Shape A)
	{
		return Tagging.A(A.Tags);
	}

	internal static bool A(Slide A)
	{
		if (Tagging.A(A.Tags))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		return Check.A(A);
	}

	internal static bool A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		return Tagging.A(A.Tags);
	}

	private static bool A(Tags A)
	{
		return Operators.CompareString(Tagging.A(A), string.Empty, TextCompare: false) != 0;
	}

	internal static void A(Tags A, string B)
	{
		string value = Tagging.A(Tagging.A(A), B);
		A.Add(Tagging.A, value);
	}

	internal static void A(SlideItem A, int B)
	{
		using List<Slide>.Enumerator enumerator = A.Slides.GetEnumerator();
		while (enumerator.MoveNext())
		{
			Tagging.A(enumerator.Current.Tags, B);
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			return;
		}
	}

	internal static void A(Tags A, int B)
	{
		Tagging.A(A, Content.XML_CURRENT_VERSION, B.ToString());
	}

	internal static void B(Tags A, int B)
	{
		string c = ((B == 0) ? "" : B.ToString());
		Tagging.A(A, Content.XML_IGNORED_VERSION, c);
	}

	private static string A(string A, string B)
	{
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(A);
		xmlDocument.DocumentElement.SelectSingleNode(Content.XML_PATH).FirstChild.InnerText = CloudStorage.AddPlaceholdersToPath(B);
		return xmlDocument.OuterXml;
	}

	internal static void A(Tags A, string B, string C)
	{
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(Tagging.A(A));
		xmlDocument.DocumentElement.SelectSingleNode(B).InnerText = C;
		A.Add(Tagging.A, xmlDocument.OuterXml);
		xmlDocument = null;
	}

	internal static void B(Tags A)
	{
		NG.A.Application.StartNewUndoEntry();
		A.Delete(Tagging.A);
	}
}
