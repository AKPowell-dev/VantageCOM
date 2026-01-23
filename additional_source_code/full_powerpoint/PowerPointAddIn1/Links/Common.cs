using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class Common
{
	[CompilerGenerated]
	internal sealed class YE
	{
		public wpfLinkRefresh A;

		public int A;

		public int B;

		public double A;

		[SpecialName]
		internal void A()
		{
			this.A.tbStatus.Text = AH.A(170780) + this.A + AH.A(93952) + B + AH.A(17804) + string.Format(AH.A(170813), this.A) + AH.A(170826);
		}
	}

	[CompilerGenerated]
	internal sealed class ZE
	{
		public wpfLinkRefresh A;

		public int A;

		public int B;

		public int C;

		[SpecialName]
		internal void A()
		{
			this.A.pbLink.Value = this.A;
			if (this.A < 100)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						this.A.tbStatus.Text = AH.A(170780) + B + AH.A(93952) + C + AH.A(17804) + string.Format(AH.A(170813), (double)this.A / 100.0) + AH.A(170826);
						return;
					}
				}
			}
			this.A.tbStatus.Text = AH.A(170847);
		}
	}

	internal static readonly string A = AH.A(94043);

	public static bool IsLinked(Tags tags)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		string text = default(string);
		bool result = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 136:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0018;
						case 4:
							goto IL_003e;
						case 5:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 6:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_003e:
					num2 = 4;
					text = tags[Base.TAG_LINK_SOURCE];
					break;
					IL_0007:
					num2 = 2;
					text = tags[Base.TAG_LINK_XML];
					goto IL_0018;
					IL_0018:
					num2 = 3;
					if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
					{
						break;
					}
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
					goto IL_003e;
					end_IL_0000_2:
					break;
				}
				num2 = 5;
				result = text.Length > 0;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 136;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static string GetAuthorFromSlide(Slide sld)
	{
		return PresentationAuthor((Microsoft.Office.Interop.PowerPoint.Presentation)sld.Parent);
	}

	public static string GetAuthorFromShape(Shape shp)
	{
		return PresentationAuthor(A(shp));
	}

	private static Microsoft.Office.Interop.PowerPoint.Presentation A(Shape A)
	{
		object obj = A;
		while (!(obj is Microsoft.Office.Interop.PowerPoint.Presentation))
		{
			obj = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj, null, AH.A(28234), new object[0], null, null, null));
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
			return (Microsoft.Office.Interop.PowerPoint.Presentation)obj;
		}
	}

	public static string PresentationAuthor(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		string result;
		try
		{
			result = Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(pres.BuiltInDocumentProperties, null, AH.A(93716), new object[1] { AH.A(93725) }, null, null, null), null, AH.A(93748), new object[0], null, null, null));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			result = "";
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static string GetLinkSource(XmlNode nd)
	{
		return CloudStorage.FillPlaceholdersInPath(nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_SOURCE).InnerText);
	}

	public static string GetLinkSourceModified(XmlNode nd)
	{
		return nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_SOURCE_LAST_MOD).InnerText;
	}

	public static string GetLinkId(XmlNode nd)
	{
		return nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_LINK_ID).InnerText;
	}

	public static string GetParentId(XmlNode nd)
	{
		return nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_PARENT_ID).InnerText;
	}

	public static string GetLinkTime(XmlNode nd)
	{
		return nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_TIME).InnerText;
	}

	public static string GetLinkUser(XmlNode nd)
	{
		return nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_USER).InnerText;
	}

	public static string GetLinkAddress(XmlNode nd)
	{
		return nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_ADDRESS).InnerText;
	}

	public static string GetLinkKeepSourceFormatting(XmlNode nd)
	{
		return nd.SelectSingleNode(AH.A(93759) + Base.XML_NODE_KEEP_SOURCE_FORMATTING).InnerText;
	}

	public static string GetLinkOther(XmlNode nd, string strNodeName)
	{
		return nd.SelectSingleNode(AH.A(93759) + strNodeName).InnerText;
	}

	public static void UpdateSource(Tags tags, RefreshInstance refreshInstance, string strFullName, bool blnUpdateLastModified)
	{
		A(tags, Base.XML_NODE_SOURCE, CloudStorage.AddPlaceholdersToPath(strFullName));
		B(tags, Base.TAG_LINK_SOURCE, strFullName);
		if (!blnUpdateLastModified)
		{
			return;
		}
		string text = ((refreshInstance == null) ? Updates.GetLastModifiedTime(strFullName) : refreshInstance.GetLastModifiedTime(strFullName));
		if (text.Length <= 0)
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
			A(tags, Base.XML_NODE_SOURCE_LAST_MOD, text);
			B(tags, Base.TAG_LINK_SOURCE_LAST_MOD, text);
			return;
		}
	}

	public static void UpdateTime(Tags tags)
	{
		string c = Base.LastUpdate();
		A(tags, Base.XML_NODE_TIME, c);
		B(tags, Base.TAG_LINK_TIME, c);
	}

	public static void UpdateUser(Tags tags, string strUser)
	{
		A(tags, Base.XML_NODE_USER, strUser);
		B(tags, Base.TAG_LINK_USER, strUser);
	}

	public static void UpdateAddress(Tags tags, string strAddress)
	{
		A(tags, Base.XML_NODE_ADDRESS, strAddress);
		B(tags, Base.TAG_LINK_ADDRESS, strAddress);
	}

	public static void UpdateType(Tags tags, ImportType type)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected I4, but got Unknown
		//IL_001a: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected I4, but got Unknown
		A(tags, Base.XML_NODE_TYPE, ((int)type).ToString());
		B(tags, Base.TAG_LINK_TYPE, ((int)type).ToString());
	}

	public static void UpdateName(Tags tags, string strName)
	{
		A(tags, Base.XML_NODE_LINK_ID, strName);
		B(tags, Base.TAG_LINK_NAME, strName);
	}

	private static void A(Tags A, string B, string C)
	{
		string tAG_LINK_XML = Base.TAG_LINK_XML;
		string text = A[tAG_LINK_XML];
		if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
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
			XmlDocument xmlDocument = UpdateXml(text, B, C);
			A.Add(tAG_LINK_XML, xmlDocument.OuterXml);
			xmlDocument = null;
			return;
		}
	}

	public static void UpdateParentId(Tags tags, string strParentId)
	{
		A(tags, Base.XML_NODE_PARENT_ID, strParentId);
	}

	public static XmlDocument UpdateXml(string strXml, string strNode, string strValue)
	{
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(strXml);
		XmlNodeList xmlNodeList = xmlDocument.SelectNodes(AH.A(93764) + strNode);
		if (xmlNodeList != null)
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = xmlNodeList.GetEnumerator();
				while (enumerator.MoveNext())
				{
					XmlNode xmlNode = (XmlNode)enumerator.Current;
					if (xmlNode.HasChildNodes && xmlNode.FirstChild.NodeType == XmlNodeType.CDATA)
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
						xmlNode.FirstChild.InnerText = strValue;
					}
					else
					{
						xmlNode.InnerText = strValue;
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_00a1;
					}
					continue;
					end_IL_00a1:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			xmlNodeList = null;
		}
		return xmlDocument;
	}

	private static void B(Tags A, string B, string C)
	{
		if (Operators.CompareString(A[B], string.Empty, TextCompare: false) == 0)
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
			A.Add(B, C);
			return;
		}
	}

	public static void AddNewFormatLinkTag(Tags tags, Link link)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		tags.Add(Base.TAG_LINK_XML, Add.GenerateXml("", link.Type, link.Source, link.SourceModified, link.LastUpdate, link.LastUser, "", link.Name, "", link.Address, (bool?)null));
	}

	public static void NavigateToSlide(object obj, Microsoft.Office.Interop.PowerPoint.Application ppApp)
	{
		object objectValue = RuntimeHelpers.GetObjectValue(obj);
		while (!(objectValue is Slide))
		{
			objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, AH.A(28234), new object[0], null, null, null));
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
			ppApp.ActiveWindow.View.GotoSlide(((Slide)objectValue).SlideIndex);
			objectValue = null;
			return;
		}
	}

	public static void UpdateProgressStart(wpfLinkRefresh frm, int intCurrent, int intTotal)
	{
		double A = (double)checked(intCurrent - 1) / (double)intTotal;
		frm.Dispatcher.Invoke([SpecialName] () =>
		{
			frm.tbStatus.Text = AH.A(170780) + intCurrent + AH.A(93952) + intTotal + AH.A(17804) + string.Format(AH.A(170813), A) + AH.A(170826);
		});
		Common.A(frm);
	}

	public static void UpdateProgressFinish(wpfLinkRefresh frm, int intCurrent, int intTotal)
	{
		int A = checked((int)Math.Round((double)intCurrent / (double)intTotal * 100.0));
		frm.Dispatcher.Invoke([SpecialName] () =>
		{
			frm.pbLink.Value = A;
			if (A < 100)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						frm.tbStatus.Text = AH.A(170780) + intCurrent + AH.A(93952) + intTotal + AH.A(17804) + string.Format(AH.A(170813), (double)A / 100.0) + AH.A(170826);
						return;
					}
				}
			}
			frm.tbStatus.Text = AH.A(170847);
		});
		Common.A(frm);
	}

	private static void A(wpfLinkRefresh A)
	{
		A.Dispatcher.Invoke(DispatcherPriority.Background, (ThreadStart)([SpecialName] () =>
		{
		}));
	}

	public static void BreakLink(Tags tags)
	{
		Tags tags2 = tags;
		try
		{
			tags2.Delete(Base.TAG_LINK_XML);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		for (int i = tags2.Count; i >= 1; i = checked(i + -1))
		{
			if (Operators.CompareString(Strings.Left(tags2.Name(i), Base.NAME_PREFIX.Length), Base.NAME_PREFIX, TextCompare: false) == 0)
			{
				tags2.Delete(tags2.Name(i));
			}
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			tags2 = null;
			return;
		}
	}

	public static void LinkWizard()
	{
		if (!Access.AllowPowerPointOperation((PlanType)5, (Restriction)2, false))
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
			if (!clsRibbon.CallbackSlideView(ShowWarning: true))
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				bool flag = false;
				try
				{
					IEnumerable<wpfManageLinks> source = System.Windows.Application.Current.Windows.OfType<wpfManageLinks>();
					if (source.Any())
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
						source.ElementAt(0).Activate();
						flag = true;
					}
					source = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				if (!flag)
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
					wpfManageLinks wpfManageLinks2 = new wpfManageLinks();
					if (Properties.ManageLinksHeight > 0.0)
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
						wpfManageLinks2.Height = Properties.ManageLinksHeight;
					}
					if (Properties.ManageLinksWidth > 0.0)
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
						wpfManageLinks2.Width = Properties.ManageLinksWidth;
					}
					wpfManageLinks2.Show();
					wpfManageLinks2 = null;
				}
				clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)10, AH.A(93779));
				return;
			}
		}
	}

	public static bool IsManageLinksDialogOpen()
	{
		try
		{
			IEnumerable<wpfManageLinks> source = System.Windows.Application.Current.Windows.OfType<wpfManageLinks>();
			if (source.Any())
			{
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
					wpfManageLinks obj = source.ElementAt(0);
					obj.Topmost = false;
					Forms.WarningMessage(AH.A(93804));
					obj.Topmost = true;
					return true;
				}
			}
			source = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return false;
	}

	public static void CloseOpenedPresentations(List<Microsoft.Office.Interop.PowerPoint.Presentation> listPresentations, bool blnSave)
	{
		using List<Microsoft.Office.Interop.PowerPoint.Presentation>.Enumerator enumerator = listPresentations.GetEnumerator();
		while (enumerator.MoveNext())
		{
			Microsoft.Office.Interop.PowerPoint.Presentation current = enumerator.Current;
			try
			{
				if (blnSave)
				{
					current.Save();
				}
				current.Close();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		while (true)
		{
			switch (5)
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

	public static void LinkEditFailed(List<string> listErrors, int intLinks)
	{
		if (listErrors == null || !listErrors.Any())
		{
			return;
		}
		if (intLinks == 1)
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
					Forms.ErrorMessage(AH.A(93913) + listErrors[0]);
					return;
				}
			}
		}
		int count = listErrors.Count;
		if (listErrors.Distinct().ToList().Count == 1)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					Forms.ErrorMessage(count + AH.A(93952) + intLinks + AH.A(93961) + listErrors[0]);
					return;
				}
			}
		}
		Forms.ErrorMessage(count + AH.A(93952) + intLinks + AH.A(94004));
	}

	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)10, strActivity);
	}
}
