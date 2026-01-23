using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane.UI;
using MacabacusMacros.Libraries.Versioning;
using Macabacus_Word.Links;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Library2.Versioning;

public sealed class Tagging
{
	internal static readonly string A = XC.A(7230);

	internal static void A(Microsoft.Office.Interop.Word.Shape A, ContentItem B, string C, string D)
	{
		if (B.Id.Length <= 0)
		{
			return;
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
			try
			{
				A.AlternativeText = Tagging.A(PC.A.Application.ActiveDocument, B, C, D);
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

	internal static void A(InlineShape A, ContentItem B, string C, string D)
	{
		if (B.Id.Length <= 0)
		{
			return;
		}
		while (true)
		{
			switch (2)
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
				A.AlternativeText = Tagging.A(PC.A.Application.ActiveDocument, B, C, D);
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

	internal static void A(Table A, ContentItem B, string C, string D)
	{
		if (B.Id.Length <= 0)
		{
			return;
		}
		while (true)
		{
			switch (1)
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
				A.Descr = Tagging.A(PC.A.Application.ActiveDocument, B, C, D);
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

	internal static string A(Document A, ContentItem B, string C, string D)
	{
		if (Tagging.A(A))
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
					return Tagging.A(Tagging.A(A.CustomXMLParts, Content.GenerateXml(B, C, D, true)));
				}
			}
		}
		return "";
	}

	private static bool A(Document A)
	{
		if (A.Path.Length > 0)
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
			if (Path.GetExtension(A.Name).Length == 5)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		if (A.Path.Length == 0)
		{
			return true;
		}
		return false;
	}

	private static string A(string A)
	{
		return XC.A(7219) + Tagging.A + XC.A(7222) + A + XC.A(7225) + Tagging.A + XC.A(7222);
	}

	private static string A(CustomXMLParts A, string B)
	{
		CustomXMLPart customXMLPart = A.Add(B, RuntimeHelpers.GetObjectValue(Missing.Value));
		string id = customXMLPart.Id;
		Marshal.ReleaseComObject(customXMLPart);
		return id;
	}

	internal static ContentInfo? A(Microsoft.Office.Interop.Word.Shape A)
	{
		return Tagging.A(A.AlternativeText);
	}

	internal static ContentInfo? A(InlineShape A)
	{
		return Tagging.A(A.AlternativeText);
	}

	internal static ContentInfo? A(Table A)
	{
		return Tagging.A(A.Descr);
	}

	private static ContentInfo? A(string A)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			return Content.GetContentInfoFromXml(Tagging.A(A));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static XmlDocument A(string A)
	{
		XmlDocument xmlDocument = new XmlDocument();
		try
		{
			xmlDocument.LoadXml(Tagging.A(PC.A.Application.ActiveDocument, A).XML);
			return xmlDocument;
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
		return null;
	}

	private static CustomXMLPart A(Document A, string B)
	{
		try
		{
			return A.CustomXMLParts.SelectByID(Tagging.B(B));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return null;
	}

	private static string B(string A)
	{
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(A);
		return xmlDocument.InnerText;
	}

	internal static bool A(Microsoft.Office.Interop.Word.Shape A)
	{
		return Tagging.A(A.AlternativeText);
	}

	internal static bool A(InlineShape A)
	{
		return Tagging.A(A.AlternativeText);
	}

	internal static bool A(Table A)
	{
		return Tagging.A(A.Descr);
	}

	internal static bool B(Document A)
	{
		bool result = default(bool);
		return result;
	}

	private static bool A(string A)
	{
		return A.Contains(Tagging.A);
	}

	internal static void A(Microsoft.Office.Interop.Word.Shape A, string B)
	{
		Tagging.A(A.AlternativeText, B);
	}

	internal static void A(InlineShape A, string B)
	{
		Tagging.A(A.AlternativeText, B);
	}

	private static void A(string A, string B)
	{
		Tagging.A(A, Content.XML_PATH, CloudStorage.AddPlaceholdersToPath(B));
	}

	public static void UpdateCurrentVersion(Microsoft.Office.Interop.Word.Shape shp, int intVersion)
	{
		A(shp.AlternativeText, Content.XML_CURRENT_VERSION, intVersion.ToString());
	}

	public static void UpdateCurrentVersion(InlineShape shp, int intVersion)
	{
		A(shp.AlternativeText, Content.XML_CURRENT_VERSION, intVersion.ToString());
	}

	public static void UpdateIgnoredVersion(Microsoft.Office.Interop.Word.Shape shp, int intVersion)
	{
		object obj;
		if (intVersion != 0)
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
			obj = intVersion.ToString();
		}
		else
		{
			obj = "";
		}
		string c = (string)obj;
		A(shp.AlternativeText, Content.XML_IGNORED_VERSION, c);
	}

	public static void UpdateIgnoredVersion(InlineShape shp, int intVersion)
	{
		object obj;
		if (intVersion != 0)
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
			obj = intVersion.ToString();
		}
		else
		{
			obj = "";
		}
		string c = (string)obj;
		A(shp.AlternativeText, Content.XML_IGNORED_VERSION, c);
	}

	internal static void A(string A, string B, string C)
	{
		Macabacus_Word.Links.CustomXML.Update(PC.A.Application.ActiveDocument, A, B, C);
	}

	public static void UnlinkContent(Microsoft.Office.Interop.Word.Shape shp)
	{
		CustomXML.RemoveCustomXMLPart(shp);
		shp.AlternativeText = "";
	}

	public static void UnlinkContent(InlineShape shp)
	{
		CustomXML.RemoveCustomXMLPart(shp);
		shp.AlternativeText = "";
	}
}
