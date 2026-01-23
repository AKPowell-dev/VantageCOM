using System;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.Links;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class CustomXML
{
	public static void Update(Document doc, string strId, string strNode, string strValue)
	{
		try
		{
			CustomXML.UpdateNode(Macabacus_Word.CustomXML.RetrievePart(doc, strId), strNode, strValue);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void UpdatePart(CustomXMLPart part, RefreshInstance refreshInstance, string strFullName, Link link, string strAddress, string strLastUpdate, string strUser, string parentId = null)
	{
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Expected I4, but got Unknown
		string text = CustomXML.XPathQuery(part);
		string text2 = "";
		CustomXMLPart customXMLPart = part;
		customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_SOURCE).FirstChild.Text = CloudStorage.AddPlaceholdersToPath(strFullName);
		try
		{
			text2 = ((refreshInstance == null) ? Updates.GetLastModifiedTime(strFullName) : refreshInstance.GetLastModifiedTime(strFullName));
			if (text2.Length > 0)
			{
				customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_SOURCE_LAST_MOD).Text = text2;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_NAME).FirstChild.Text = link.Name;
		customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_TYPE).Text = ((int)link.Type).ToString();
		customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_UPDATED).Text = strLastUpdate;
		try
		{
			customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_USER).Text = strUser;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		try
		{
			customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_ADDRESS).FirstChild.Text = strAddress;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		if (parentId != null)
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
			try
			{
				customXMLPart.SelectSingleNode(text + CustomXML.XML_NODE_PARENT).Text = parentId;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		customXMLPart = null;
	}

	public static XmlDocument GetLinkXML(Document doc, string strId)
	{
		XmlDocument xmlDocument = new XmlDocument();
		try
		{
			xmlDocument.LoadXml(Macabacus_Word.CustomXML.RetrievePart(doc, strId).XML);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			xmlDocument = null;
			ProjectData.ClearProjectError();
		}
		return xmlDocument;
	}
}
