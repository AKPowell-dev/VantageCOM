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
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2.Versioning;

public sealed class Tagging
{
	internal static readonly string A = VH.A(84124);

	internal static void A(Shape A, ContentItem B, string C, string D)
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
				A.AlternativeText = Tagging.A(MH.A.Application.ActiveWorkbook, B, C, D);
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

	internal static string A(Microsoft.Office.Interop.Excel.Workbook A, ContentItem B, string C, string D)
	{
		if (Tagging.A(A))
		{
			while (true)
			{
				switch (5)
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

	private static bool A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		if (A.Path.Length > 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Path.GetExtension(A.Name).Length == 5)
			{
				while (true)
				{
					switch (5)
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
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		return false;
	}

	private static string A(string A)
	{
		return VH.A(75525) + Tagging.A + VH.A(84116) + A + VH.A(84119) + Tagging.A + VH.A(84116);
	}

	private static string A(CustomXMLParts A, string B)
	{
		CustomXMLPart customXMLPart = A.Add(B, RuntimeHelpers.GetObjectValue(Missing.Value));
		string id = customXMLPart.Id;
		Marshal.ReleaseComObject(customXMLPart);
		return id;
	}

	internal static ContentInfo? A(Shape A)
	{
		return Tagging.A(A.AlternativeText);
	}

	private static ContentInfo? A(string A)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
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
			xmlDocument.LoadXml(Tagging.A(A).XML);
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

	internal static bool A(Shape A)
	{
		return Tagging.A(A.AlternativeText);
	}

	private static bool A(string A)
	{
		return A.Contains(Tagging.A);
	}

	internal static void A(Shape A, string B)
	{
		Tagging.A(A.AlternativeText, B);
	}

	private static void A(string A, string B)
	{
		Tagging.A(A, Content.XML_PATH, CloudStorage.AddPlaceholdersToPath(B));
	}

	public static void UpdateCurrentVersion(Shape shp, int intVersion)
	{
		A(shp.AlternativeText, Content.XML_CURRENT_VERSION, intVersion.ToString());
	}

	public static void UpdateIgnoredVersion(Shape shp, int intVersion)
	{
		object obj;
		if (intVersion != 0)
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
		try
		{
			CustomXML.UpdateNode(Tagging.A(A), B, C);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void UnlinkContent(Shape shp)
	{
		CustomXML.A(shp);
		shp.AlternativeText = "";
	}

	private static CustomXMLPart A(string A)
	{
		return CustomXML.A(MH.A.Application.ActiveWorkbook, A);
	}
}
