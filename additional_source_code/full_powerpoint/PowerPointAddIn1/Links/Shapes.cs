using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class Shapes
{
	public struct EditedShapes
	{
		[Serializable]
		[CompilerGenerated]
		internal sealed class _Closure_0024__
		{
			public static readonly _Closure_0024__ A;

			public static AB<object> A;

			static _Closure_0024__()
			{
				_Closure_0024__.A = new _Closure_0024__();
			}

			[SpecialName]
			internal void A(object A)
			{
				//IL_001c: Unknown result type (might be due to invalid IL or missing references)
				if (!(A is TextLink))
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					((TextLink)A).ClearReferences();
					return;
				}
			}
		}

		public List<object> Objects;

		public List<bool> IsError;

		public List<string> Errors;

		public void ClearReferences(bool doDeepClearing, bool collectGarbage = false)
		{
			ref List<object> objects = ref Objects;
			object obj;
			if (!doDeepClearing)
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
				obj = null;
			}
			else if (_Closure_0024__.A == null)
			{
				obj = (_Closure_0024__.A = [SpecialName] (object A) =>
				{
					//IL_001c: Unknown result type (might be due to invalid IL or missing references)
					if (A is TextLink)
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
								((TextLink)A).ClearReferences();
								return;
							}
						}
					}
				});
			}
			else
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
				obj = _Closure_0024__.A;
			}
			AB<object> aB = (AB<object>)obj;
			object obj2;
			if (aB != null)
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
				obj2 = new Action<object>(aB.Invoke);
			}
			else
			{
				obj2 = null;
			}
			ReleaseHelper.ClearListReferences<object>(ref objects, collectGarbage, (Action<object>)obj2);
		}
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Action<TextLink> A;

		public static Action<TextLink> B;

		public static Func<TextLink, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal void A(TextLink A)
		{
			A.ClearReferences();
		}

		[SpecialName]
		internal void B(TextLink A)
		{
			A.ClearReferences();
		}

		[SpecialName]
		internal int A(TextLink A)
		{
			return A.TextRange.Start;
		}
	}

	public static Link LinkDetails(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0287: Unknown result type (might be due to invalid IL or missing references)
		//IL_0289: Unknown result type (might be due to invalid IL or missing references)
		//IL_0214: Unknown result type (might be due to invalid IL or missing references)
		//IL_036a: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Link val = default(Link);
		string text = default(string);
		XmlDocument xmlDocument = default(XmlDocument);
		IEnumerator enumerator = default(IEnumerator);
		XmlNode nd = default(XmlNode);
		Link result = default(Link);
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
				case 803:
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
							goto IL_0011;
						case 4:
							goto IL_0029;
						case 5:
							goto IL_0050;
						case 6:
							goto IL_0059;
						case 7:
							goto IL_0064;
						case 8:
							goto IL_0097;
						case 9:
							goto IL_0099;
						case 10:
							goto IL_00ac;
						case 11:
							goto IL_00bf;
						case 12:
							goto IL_00d2;
						case 13:
							goto IL_00e5;
						case 14:
							goto IL_0102;
						case 15:
							goto IL_0115;
						case 16:
							goto IL_0128;
						case 17:
						case 18:
							goto IL_0139;
						case 19:
							goto IL_014a;
						case 20:
							goto IL_016c;
						case 22:
							goto IL_0174;
						case 23:
							goto IL_0177;
						case 24:
							goto IL_019a;
						case 25:
							goto IL_01bb;
						case 26:
							goto IL_01e0;
						case 27:
							goto IL_01ef;
						case 28:
							goto IL_0219;
						case 29:
							goto IL_023c;
						case 30:
							goto IL_025f;
						case 21:
						case 31:
						case 32:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 33:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_019a:
					num2 = 24;
					val.SourceModified = shp.Tags[Base.TAG_LINK_SOURCE_LAST_MOD].ToString();
					goto IL_01bb;
					IL_0007:
					num2 = 2;
					val = default(Link);
					goto IL_0011;
					IL_0011:
					num2 = 3;
					text = shp.Tags[Base.TAG_LINK_XML];
					goto IL_0029;
					IL_0029:
					num2 = 4;
					if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
						goto IL_0050;
					}
					goto IL_0174;
					IL_01bb:
					num2 = 25;
					val.Name = shp.Tags[Base.TAG_LINK_NAME].ToString();
					goto IL_01e0;
					IL_01e0:
					num2 = 26;
					val.ParentId = "";
					goto IL_01ef;
					IL_01ef:
					num2 = 27;
					val.Type = (ImportType)Conversions.ToInteger(shp.Tags[Base.TAG_LINK_TYPE].ToString());
					goto IL_0219;
					IL_0050:
					num2 = 5;
					xmlDocument = new XmlDocument();
					goto IL_0059;
					IL_0059:
					num2 = 6;
					xmlDocument.LoadXml(text);
					goto IL_0064;
					IL_0064:
					num2 = 7;
					enumerator = xmlDocument.SelectNodes(AH.A(94272)).GetEnumerator();
					goto IL_013c;
					IL_013c:
					if (enumerator.MoveNext())
					{
						nd = (XmlNode)enumerator.Current;
						goto IL_0097;
					}
					goto IL_014a;
					IL_014a:
					num2 = 19;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_016c;
					IL_023c:
					num2 = 29;
					val.LastUser = shp.Tags[Base.TAG_LINK_USER].ToString();
					goto IL_025f;
					IL_025f:
					num2 = 30;
					val.Address = shp.Tags[Base.TAG_LINK_ADDRESS].ToString();
					break;
					IL_016c:
					xmlDocument = null;
					break;
					IL_0219:
					num2 = 28;
					val.LastUpdate = shp.Tags[Base.TAG_LINK_TIME].ToString();
					goto IL_023c;
					IL_0097:
					num2 = 8;
					goto IL_0099;
					IL_0099:
					num2 = 9;
					val.Source = Common.GetLinkSource(nd);
					goto IL_00ac;
					IL_00ac:
					num2 = 10;
					val.SourceModified = Common.GetLinkSourceModified(nd);
					goto IL_00bf;
					IL_00bf:
					num2 = 11;
					val.Name = Common.GetLinkId(nd);
					goto IL_00d2;
					IL_00d2:
					num2 = 12;
					val.ParentId = Common.GetParentId(nd);
					goto IL_00e5;
					IL_00e5:
					num2 = 13;
					val.Type = (ImportType)Conversions.ToInteger(Common.GetLinkOther(nd, Base.XML_NODE_TYPE));
					goto IL_0102;
					IL_0102:
					num2 = 14;
					val.LastUpdate = Common.GetLinkTime(nd);
					goto IL_0115;
					IL_0115:
					num2 = 15;
					val.LastUser = Common.GetLinkUser(nd);
					goto IL_0128;
					IL_0128:
					num2 = 16;
					val.Address = Common.GetLinkAddress(nd);
					goto IL_0139;
					IL_0139:
					num2 = 18;
					goto IL_013c;
					IL_0174:
					num2 = 22;
					goto IL_0177;
					IL_0177:
					num2 = 23;
					val.Source = shp.Tags[Base.TAG_LINK_SOURCE].ToString();
					goto IL_019a;
					end_IL_0000_2:
					break;
				}
				num2 = 32;
				result = val;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 803;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool IsLinked(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return Common.IsLinked(shp.Tags);
	}

	public static void EditLinks(Selection sel)
	{
		EditedShapes editedShapes = default(EditedShapes);
		if (Common.IsManageLinksDialogOpen())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			List<object> B = new List<object>();
			if (!sel.HasChildShapeRange)
			{
				{
					enumerator = sel.ShapeRange.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B);
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00ba;
							}
							continue;
							end_IL_00ba:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			else
			{
				try
				{
					enumerator2 = sel.ChildShapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, ref B);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0063;
						}
						continue;
						end_IL_0063:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			foreach (Hyperlink item in Hyperlinks.SelectedLinks(sel))
			{
				B.Add(item);
			}
			if (B.Any())
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
				editedShapes = EditLink(B);
				Common.LogActivity(AH.A(98769));
			}
			else
			{
				Forms.WarningMessage(AH.A(95753));
			}
			editedShapes.ClearReferences(doDeepClearing: true, collectGarbage: true);
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref List<object> B)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					if (IsLinked(A))
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
						B.Add(A);
					}
					B.AddRange(Text.SelectedLinks(A).ToArray());
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Shapes.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	public static EditedShapes EditLink(List<object> listShapes)
	{
		EditedShapes result = new EditedShapes
		{
			Objects = listShapes,
			IsError = null,
			Errors = null
		};
		wpfLinkEdit wpfLinkEdit2;
		wpfLinkEdit obj = (wpfLinkEdit2 = new wpfLinkEdit(listShapes));
		if (Properties.EditLinksWidth > 0.0)
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
			wpfLinkEdit2.Width = Properties.EditLinksWidth;
		}
		Base.ShowDialogNotTopmost((System.Windows.Window)obj);
		if (wpfLinkEdit2.DialogResult == true)
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
			result = wpfLinkEdit2.ReturnValue;
		}
		wpfLinkEdit2 = null;
		GC.Collect();
		Common.LinkEditFailed(result.Errors, listShapes.Count);
		return result;
	}

	public static void RefreshLinks(Selection sel)
	{
		if (!Access.AllowSuiteOperation((PlanType)5, (Restriction)1, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			List<Microsoft.Office.Interop.PowerPoint.Shape> B = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
			List<TextLink> C = new List<TextLink>();
			List<Hyperlink> c = null;
			try
			{
				if (sel.HasChildShapeRange)
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
					try
					{
						enumerator = sel.ChildShapeRange.GetEnumerator();
						while (enumerator.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B, ref C);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_0074;
							}
							continue;
							end_IL_0074:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				else
				{
					try
					{
						enumerator2 = sel.ShapeRange.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, ref B, ref C);
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				c = Hyperlinks.SelectedLinks(sel);
				A(B, C, c, D: false);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			ReleaseHelper.ClearListReferences<Microsoft.Office.Interop.PowerPoint.Shape>(ref B, false, (Action<Microsoft.Office.Interop.PowerPoint.Shape>)null);
			ReleaseHelper.ClearListReferences<TextLink>(ref C, false, (Action<TextLink>)([SpecialName] (TextLink A) =>
			{
				A.ClearReferences();
			}));
			ReleaseHelper.ClearListReferences<Hyperlink>(ref c, false, (Action<Hyperlink>)null);
			ReleaseHelper.DoGarbageCollection();
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref List<Microsoft.Office.Interop.PowerPoint.Shape> B, ref List<TextLink> C)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			if (IsLinked(A))
			{
				B.Add(A);
			}
			C.AddRange(Text.SelectedLinks(A).ToArray());
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Shapes.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B, ref C);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	public static void UpdateAllLinks()
	{
		if (!Access.AllowSuiteOperation((PlanType)5, (Restriction)2, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
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
			bool flag = false;
			bool flag2 = false;
			List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
			List<TextLink> listTextLinks = new List<TextLink>();
			List<Hyperlink> list2 = new List<Hyperlink>();
			enumerator = NG.A.Application.ActivePresentation.Slides.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					try
					{
						enumerator2 = slide.Shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							if (shape.Visible == MsoTriState.msoTrue)
							{
								if (IsLinked(shape))
								{
									list.Add(shape);
								}
								listTextLinks.AddRange(Text.SelectedLinks(shape).ToArray());
							}
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
					if (listTextLinks.Count > 0)
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
							flag2 = Hyperlinks.PromptToConvert();
							flag = true;
						}
						if (flag2)
						{
							Hyperlinks.ConvertFromLegacyLinks(ref listTextLinks);
						}
					}
					try
					{
						enumerator3 = slide.Hyperlinks.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							Hyperlink hyperlink = (Hyperlink)enumerator3.Current;
							if (!Hyperlinks.IsLinked(hyperlink))
							{
								continue;
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
							list2.Add(hyperlink);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_016f;
							}
							continue;
							end_IL_016f:
							break;
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								(enumerator3 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_01a9;
					}
					continue;
					end_IL_01a9:
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			A(list, listTextLinks, list2, D: true);
			ReleaseHelper.ClearListReferences<Microsoft.Office.Interop.PowerPoint.Shape>(ref list, false, (Action<Microsoft.Office.Interop.PowerPoint.Shape>)null);
			Action<TextLink> action;
			if (_Closure_0024__.B == null)
			{
				action = (_Closure_0024__.B = [SpecialName] (TextLink A) =>
				{
					A.ClearReferences();
				});
			}
			else
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
				action = _Closure_0024__.B;
			}
			ReleaseHelper.ClearListReferences<TextLink>(ref listTextLinks, false, action);
			ReleaseHelper.ClearListReferences<Hyperlink>(ref list2, false, (Action<Hyperlink>)null);
			ReleaseHelper.DoGarbageCollection();
			return;
		}
	}

	private static void A(List<Microsoft.Office.Interop.PowerPoint.Shape> A, List<TextLink> B, List<Hyperlink> C, bool D)
	{
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0081: Expected O, but got Unknown
		//IL_0083: Expected O, but got Unknown
		//IL_012c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0132: Expected O, but got Unknown
		//IL_0134: Expected O, but got Unknown
		//IL_0254: Unknown result type (might be due to invalid IL or missing references)
		//IL_025a: Expected O, but got Unknown
		//IL_025c: Expected O, but got Unknown
		//IL_033f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0345: Expected O, but got Unknown
		//IL_0347: Expected O, but got Unknown
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c3: Expected O, but got Unknown
		//IL_0695: Unknown result type (might be due to invalid IL or missing references)
		Dictionary<object, string> dictionary = new Dictionary<object, string>();
		checked
		{
			int num = A.Count + B.Count + C.Count;
			int num2 = 0;
			bool flag = false;
			Microsoft.Office.Interop.PowerPoint.Application application;
			List<string> listUpdatedShapeNames;
			if (num > 0)
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
				wpfLinkRefresh wpfLinkRefresh = new wpfLinkRefresh();
				application = NG.A.Application;
				RefreshInstance refreshInstance;
				try
				{
					refreshInstance = new RefreshInstance(System.Windows.Window.GetWindow(wpfLinkRefresh));
					refreshInstance.LegalFonts = BrandCompliance.GetLegalFontTypes(application.ActivePresentation);
				}
				catch (UpdateLinkException ex)
				{
					ProjectData.SetProjectError((Exception)ex);
					UpdateLinkException ex2 = ex;
					wpfLinkRefresh = null;
					dictionary = null;
					application = null;
					Forms.WarningMessage(((Exception)(object)ex2).Message);
					ProjectData.ClearProjectError();
					return;
				}
				wpfLinkRefresh.Show();
				listUpdatedShapeNames = new List<string>();
				application.StartNewUndoEntry();
				TimelineRestorer timelineRestorer = new TimelineRestorer();
				try
				{
					CopierAsPicture copierAsPic = new CopierAsPicture();
					foreach (Microsoft.Office.Interop.PowerPoint.Shape item in A)
					{
						if (wpfLinkRefresh.Canceled)
						{
							break;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							if (refreshInstance.Canceled)
							{
								break;
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								num2++;
								Common.UpdateProgressStart(wpfLinkRefresh, num2, num);
								try
								{
									Refresh(item, D, ref listUpdatedShapeNames, ref refreshInstance, copierAsPic, timelineRestorer);
								}
								catch (UpdateLinkException ex3)
								{
									ProjectData.SetProjectError((Exception)ex3);
									UpdateLinkException ex4 = ex3;
									dictionary.Add(item, ((Exception)(object)ex4).Message);
									ProjectData.ClearProjectError();
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									dictionary.Add(item, ex6.Message);
									if (!D)
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
										clsReporting.LogException(ex6);
									}
									ProjectData.ClearProjectError();
								}
								Common.UpdateProgressFinish(wpfLinkRefresh, num2, num);
								break;
							}
							goto IL_018b;
						}
						break;
						IL_018b:;
					}
					timelineRestorer.A();
					List<TextLink> source = B;
					Func<TextLink, int> keySelector;
					if (_Closure_0024__.A == null)
					{
						keySelector = (_Closure_0024__.A = [SpecialName] (TextLink val) => val.TextRange.Start);
					}
					else
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
						keySelector = _Closure_0024__.A;
					}
					B = source.OrderBy(keySelector).ToList();
					using (List<TextLink>.Enumerator enumerator2 = B.GetEnumerator())
					{
						while (true)
						{
							IL_02b1:
							if (enumerator2.MoveNext())
							{
								TextLink current2 = enumerator2.Current;
								if (wpfLinkRefresh.Canceled)
								{
									break;
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									if (refreshInstance.Canceled)
									{
										break;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										num2++;
										Common.UpdateProgressStart(wpfLinkRefresh, num2, num);
										try
										{
											Text.Refresh(current2, D, ref listUpdatedShapeNames, ref refreshInstance);
										}
										catch (UpdateLinkException ex7)
										{
											ProjectData.SetProjectError((Exception)ex7);
											UpdateLinkException ex8 = ex7;
											dictionary.Add(current2, ((Exception)(object)ex8).Message);
											ProjectData.ClearProjectError();
										}
										catch (Exception ex9)
										{
											ProjectData.SetProjectError(ex9);
											Exception ex10 = ex9;
											dictionary.Add(current2, ex10.Message);
											if (!D)
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
												clsReporting.LogException(ex10);
											}
											ProjectData.ClearProjectError();
										}
										Common.UpdateProgressFinish(wpfLinkRefresh, num2, num);
										break;
									}
									goto IL_02b1;
								}
								break;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_02bd;
								}
								continue;
								end_IL_02bd:
								break;
							}
							break;
						}
					}
					foreach (Hyperlink item2 in C)
					{
						if (wpfLinkRefresh.Canceled)
						{
							break;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							if (refreshInstance.Canceled)
							{
								break;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								num2++;
								Common.UpdateProgressStart(wpfLinkRefresh, num2, num);
								try
								{
									Hyperlinks.Refresh(item2, D, ref listUpdatedShapeNames, ref refreshInstance);
								}
								catch (UpdateLinkException ex11)
								{
									ProjectData.SetProjectError((Exception)ex11);
									UpdateLinkException ex12 = ex11;
									dictionary.Add(item2, ((Exception)(object)ex12).Message);
									ProjectData.ClearProjectError();
								}
								catch (Exception ex13)
								{
									ProjectData.SetProjectError(ex13);
									Exception ex14 = ex13;
									dictionary.Add(item2, ex14.Message);
									if (!D)
									{
										clsReporting.LogException(ex14);
									}
									ProjectData.ClearProjectError();
								}
								Common.UpdateProgressFinish(wpfLinkRefresh, num2, num);
								break;
							}
							goto IL_0392;
						}
						break;
						IL_0392:;
					}
					Thread.Sleep(500);
					int num3;
					if (!wpfLinkRefresh.Canceled)
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
						if (!refreshInstance.Canceled)
						{
							num3 = 0;
							goto IL_03e8;
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
					num3 = ((num2 < num) ? 1 : 0);
					goto IL_03e8;
					IL_03e8:
					flag = unchecked((byte)num3) != 0;
					wpfLinkRefresh.Close();
					wpfLinkRefresh = null;
					if (!D)
					{
						try
						{
							listUpdatedShapeNames = listUpdatedShapeNames.Distinct().ToList();
							if (listUpdatedShapeNames.Count > 0)
							{
								application.ActiveWindow.Selection.SlideRange[1].Shapes.Range(listUpdatedShapeNames.ToArray()).Select();
							}
						}
						catch (Exception ex15)
						{
							ProjectData.SetProjectError(ex15);
							Exception ex16 = ex15;
							ProjectData.ClearProjectError();
						}
					}
				}
				catch (Exception ex17)
				{
					ProjectData.SetProjectError(ex17);
					Exception ex18 = ex17;
					if (wpfLinkRefresh != null)
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
						try
						{
							wpfLinkRefresh.Close();
						}
						catch (Exception projectError)
						{
							ProjectData.SetProjectError(projectError);
							ProjectData.ClearProjectError();
						}
						wpfLinkRefresh = null;
					}
					clsReporting.LogException(ex18);
					ProjectData.ClearProjectError();
				}
				ExcelToPowerPoint.ActivatePowerPoint(application);
				Base.ReleaseRefreshInstance(ref refreshInstance, true);
				if (flag)
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
					System.Windows.Forms.MessageBox.Show(AH.A(98800) + num2 + AH.A(93952) + num + AH.A(98873), AH.A(5874), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
			}
			if (!dictionary.Any())
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
				if (num > 0)
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
					if (D)
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
						if (!flag)
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
							Forms.SuccessMessage(AH.A(98888));
						}
					}
				}
				else if (!D)
				{
					Forms.InfoMessage(AH.A(98977));
				}
				else
				{
					Forms.InfoMessage(AH.A(99040));
				}
			}
			else if (num == 1)
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
				Forms.ErrorMessage(dictionary.Values.ElementAtOrDefault(0));
			}
			else if (System.Windows.Forms.MessageBox.Show(AH.A(99135), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Hand) == DialogResult.Yes)
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
				List<LinkError> list = new List<LinkError>();
				foreach (KeyValuePair<object, string> item3 in dictionary)
				{
					string text = AH.A(99287);
					string strName;
					try
					{
						if (item3.Key is Microsoft.Office.Interop.PowerPoint.Shape)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								text = AH.A(99316);
								strName = ((Microsoft.Office.Interop.PowerPoint.Shape)item3.Key).Name;
								break;
							}
						}
						else if (item3.Key is TextLink)
						{
							text = AH.A(99347);
							strName = ((TextLink)item3.Key).TextRange.Text;
						}
						else
						{
							text = AH.A(99386);
							strName = ((Hyperlink)item3.Key).ScreenTip;
						}
					}
					catch (Exception projectError2)
					{
						ProjectData.SetProjectError(projectError2);
						strName = text;
						ProjectData.ClearProjectError();
					}
					list.Add(new LinkError(RuntimeHelpers.GetObjectValue(item3.Key), strName, item3.Value));
				}
				wpfLinkUpdateErrors obj = new wpfLinkUpdateErrors();
				obj.colShapeOrSlide.Header = AH.A(99425);
				obj.lvErrors.ItemsSource = list;
				obj.Show();
				_ = null;
				list = null;
			}
			application = null;
			listUpdatedShapeNames = null;
			dictionary = null;
			Common.LogActivity(AH.A(99436));
		}
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Refresh(Microsoft.Office.Interop.PowerPoint.Shape shp, bool blnAll, ref List<string> listUpdatedShapeNames, ref RefreshInstance refreshInstance, CopierAsPicture copierAsPic, TimelineRestorer restorer = null)
	{
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
		if (blnAll)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Common.NavigateToSlide(shp, NG.A.Application);
		}
		shp.Select();
		int id = shp.Id;
		restorer?.A(shp);
		A(ref shp, ref refreshInstance, LinkDetails(shp), copierAsPic);
		try
		{
			shape = NG.A.Application.ActiveWindow.Selection.ShapeRange[1];
			if (restorer != null)
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
				restorer.A(shape, id);
			}
			listUpdatedShapeNames.Add(shape.Name);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return shape;
	}

	private static void A(ref Microsoft.Office.Interop.PowerPoint.Shape A, ref RefreshInstance B, Link C, CopierAsPicture D)
	{
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ac: Invalid comparison between Unknown and I4
		//IL_01cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d3: Invalid comparison between Unknown and I4
		//IL_01ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0206: Unknown result type (might be due to invalid IL or missing references)
		//IL_0208: Unknown result type (might be due to invalid IL or missing references)
		//IL_020d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0230: Unknown result type (might be due to invalid IL or missing references)
		//IL_0235: Unknown result type (might be due to invalid IL or missing references)
		//IL_0237: Unknown result type (might be due to invalid IL or missing references)
		//IL_023f: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_01de: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0300: Unknown result type (might be due to invalid IL or missing references)
		//IL_0302: Unknown result type (might be due to invalid IL or missing references)
		//IL_0307: Unknown result type (might be due to invalid IL or missing references)
		//IL_0309: Unknown result type (might be due to invalid IL or missing references)
		//IL_030c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0342: Expected I4, but got Unknown
		//IL_00ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0347: Unknown result type (might be due to invalid IL or missing references)
		//IL_05fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0682: Unknown result type (might be due to invalid IL or missing references)
		//IL_0684: Unknown result type (might be due to invalid IL or missing references)
		//IL_0689: Unknown result type (might be due to invalid IL or missing references)
		//IL_068b: Unknown result type (might be due to invalid IL or missing references)
		//IL_06b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_06cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_062a: Unknown result type (might be due to invalid IL or missing references)
		//IL_06a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0126: Unknown result type (might be due to invalid IL or missing references)
		//IL_0907: Unknown result type (might be due to invalid IL or missing references)
		//IL_0374: Unknown result type (might be due to invalid IL or missing references)
		//IL_0376: Unknown result type (might be due to invalid IL or missing references)
		//IL_0384: Unknown result type (might be due to invalid IL or missing references)
		//IL_07a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_07aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_07af: Unknown result type (might be due to invalid IL or missing references)
		//IL_08f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0651: Unknown result type (might be due to invalid IL or missing references)
		//IL_078e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0790: Unknown result type (might be due to invalid IL or missing references)
		//IL_0795: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		Range range = null;
		Microsoft.Office.Interop.Excel.Chart chart = null;
		Name name = null;
		Worksheet worksheet = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		Regex regex = new Regex(AH.A(99473));
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		string text = "";
		string text2 = "";
		bool flag4 = false;
		bool flag5 = Base.SourceIsRange(C);
		B.LocateSource(ref workbook, ref range, ref chart, ref C, ref flag, ref flag2, flag5);
		if (flag2)
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
			Common.UpdateSource(A.Tags, B, C.Source, blnUpdateLastModified: false);
		}
		if (flag)
		{
			return;
		}
		XlSheetVisibility xlSheetVisibility = default(XlSheetVisibility);
		PasteParameters val3 = default(PasteParameters);
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = default(Microsoft.Office.Interop.PowerPoint.Shape);
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (range == null)
			{
				if (chart == null)
				{
					if (flag5)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								workbook = null;
								throw new UpdateLinkException(AH.A(94838));
							}
						}
					}
					workbook = null;
					throw new UpdateLinkException(AH.A(99593));
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			Link val = C;
			Link val2 = LinkDetails(A);
			if (flag5)
			{
				B.SourceRange(val2, workbook, ref name, ref worksheet, ref range, ref text2, ref xlSheetVisibility);
			}
			else
			{
				if (A.HasChart == MsoTriState.msoTrue)
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
					if (Operators.ConditionalCompareObjectEqual(A.Chart.get_HasAxis((object)Microsoft.Office.Core.XlAxisType.xlValue, RuntimeHelpers.GetObjectValue(Missing.Value)), MsoTriState.msoTrue, TextCompare: false))
					{
						text = Shapes.A(A);
					}
				}
				B.SourceChart(val2, workbook, ref chart);
				flag3 = true;
			}
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			Slide targetSlide = Shapes.A(A, application);
			bool flag6;
			try
			{
				A.PickupAnimation();
				flag6 = true;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				flag6 = false;
				ProjectData.ClearProjectError();
			}
			try
			{
				flag4 = Conversions.ToBoolean(NewLateBinding.LateGet(A, null, AH.A(69417), new object[0], null, null, null));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			if ((int)val2.Type == 5)
			{
				if (val2.Type == val.Type)
				{
					goto IL_02f7;
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
			if ((int)val2.Type == 4)
			{
				if (val2.Type == val.Type)
				{
					goto IL_02f7;
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
			shape2 = ExcelToPowerPoint.DestinationShape(A, val2.Type);
			val3.PasteType = val2.Type;
			val3.MatchSize = clsImportExport.GetMatchSize(PB.Settings.ImportMatchDestinationWidth, PB.Settings.ImportMatchDestinationHeight);
			val3.CallingApplication = (CallingApp)2;
			val3.SourceRange = range;
			val3.SourceChart = chart;
			val3.TargetApplication = application;
			val3.TargetSlide = targetSlide;
			val3.TargetShape = shape2;
			val3.TargetWidth = shape2.Width;
			val3.TargetHeight = shape2.Height;
			val3.TargetTop = shape2.Top;
			val3.TargetLeft = shape2.Left;
			val3.TargetZOrder = shape2.ZOrderPosition;
			val3.TargetShadow = shape2.Shadow;
			val3.TargetReflection = shape2.Reflection;
			val3.TargetGlow = shape2.Glow;
			val3.TargetThreeD = shape2.ThreeD;
			goto IL_02f7;
			IL_02f7:
			B.CheckCalculationMode(val2);
			ImportType type = val2.Type;
			checked
			{
				switch (unchecked(type - 1))
				{
				case 0:
				case 10:
					shape = CellsAsPowerPointPicture.Paste(val3, ref B, D, true);
					break;
				case 1:
				{
					if (KG.A.RebuildTables)
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
						shape2 = ExcelToPowerPoint.PrepPlaceholder(shape2, val2.Type);
						shape = CellsAsPowerPointTable.Paste(val3);
						break;
					}
					int num = Shapes.A((Range)range.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]);
					Rows rows = A.Table.Rows;
					while (rows.Count < num)
					{
						try
						{
							rows.Add(rows.Count - 1);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							rows.Add();
							ProjectData.ClearProjectError();
						}
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
					while (rows.Count > num)
					{
						rows[rows.Count - 1].Delete();
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
					rows = null;
					int num2 = Shapes.A((Range)range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]);
					Columns columns = A.Table.Columns;
					while (columns.Count < num2)
					{
						try
						{
							columns.Add(columns.Count - 1);
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							columns.Add();
							ProjectData.ClearProjectError();
						}
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
					while (columns.Count > num2)
					{
						columns[columns.Count - 1].Delete();
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
					columns = null;
					range.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
					Microsoft.Office.Interop.PowerPoint.Shape shape3 = application.ActiveWindow.Selection.SlideRange[1].Shapes.Paste()[1];
					int num3 = num;
					for (int i = 1; i <= num3; i++)
					{
						int num4 = num2;
						for (int j = 1; j <= num4; j++)
						{
							try
							{
								A.Table.Cell(i, j).Shape.TextFrame2.TextRange.Text = shape3.Table.Cell(i, j).Shape.TextFrame2.TextRange.Text;
							}
							catch (Exception ex9)
							{
								ProjectData.SetProjectError(ex9);
								Exception ex10 = ex9;
								ProjectData.ClearProjectError();
							}
						}
					}
					shape3.Delete();
					shape3 = null;
					break;
				}
				case 2:
					shape = CellsAsPowerPointEmbedded.Paste(val3, ref B);
					break;
				case 3:
					if (range.Cells.Count > 1)
					{
						throw new UpdateLinkException(AH.A(94708));
					}
					if (A.HasTextFrame == MsoTriState.msoFalse)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								throw new UpdateLinkException(AH.A(99500));
							}
						}
					}
					A.TextFrame.TextRange.Text = range.Text.ToString().Trim();
					break;
				case 4:
					if (val2.Type == val.Type)
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
						CellsAsPowerPointChart.PopulateChartData(range, A);
					}
					else
					{
						shape = CellsAsPowerPointChart.Paste(val3);
					}
					break;
				case 5:
				case 11:
					shape = ChartAsPowerPointPicture.Paste(val3, D, true);
					break;
				case 6:
					shape = ChartAsPowerPointChart.Paste(val3);
					break;
				case 7:
					shape = ChartAsPowerPointEmbedded.Paste(val3);
					break;
				}
				if (shape == null)
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
					ConvertLegacyLink(A);
					Common.UpdateSource(A.Tags, B, workbook.FullName, blnUpdateLastModified: true);
					Common.UpdateUser(A.Tags, workbook.Application.UserName);
					Common.UpdateTime(A.Tags);
					if (name != null)
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
						Common.UpdateAddress(A.Tags, name.RefersTo.ToString());
					}
					if (A != null)
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
						A.Select();
					}
				}
				else
				{
					if (flag3)
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
						Add.LinkToChartPowerPoint(shape, ref chart, B, val2.Type, val2.ParentId);
					}
					else
					{
						Add.LinkToRangePowerPoint(shape, range, B, val2.Type, val2.ParentId, name);
					}
					if (flag6)
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
						try
						{
							shape.ApplyAnimation();
						}
						catch (Exception ex11)
						{
							ProjectData.SetProjectError(ex11);
							Exception ex12 = ex11;
							ProjectData.ClearProjectError();
						}
					}
					if (flag4)
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
						try
						{
							NewLateBinding.LateSetComplex(shape, null, AH.A(69417), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
						}
						catch (Exception ex13)
						{
							ProjectData.SetProjectError(ex13);
							Exception ex14 = ex13;
							ProjectData.ClearProjectError();
						}
					}
					if (text.Length > 0)
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
						if (shape.HasChart == MsoTriState.msoTrue)
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
							if (regex.IsMatch(text))
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
								Shapes.A(shape, text);
								clsCharts.RescaleAxis(shape);
							}
							regex = null;
						}
					}
				}
				if (worksheet != null)
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
					if (worksheet.Visible != xlSheetVisibility)
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
						worksheet.Visible = xlSheetVisibility;
					}
				}
				B.RestoreExcel();
				RefreshInstance val4;
				(val4 = B).RefreshedLinkCount = val4.RefreshedLinkCount + 1;
				application = null;
				targetSlide = null;
				A = null;
				worksheet = null;
				name = null;
				range = null;
				chart = null;
				workbook = null;
				return;
			}
		}
	}

	private static int A(Range A)
	{
		if (A.Cells.Count > 1)
		{
			return A.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value)).Count;
		}
		if (Operators.ConditionalCompareObjectEqual(A.Hidden, false, TextCompare: false))
		{
			return 1;
		}
		return 0;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		bool flag = false;
		Table table = A.Table;
		if (Operators.CompareString(table.Style.Id, AH.A(99716), TextCompare: false) != 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (table.Style.Id.Length != 0)
			{
				B.Table.ApplyStyle(table.Style.Id);
				B.Table.FirstRow = table.FirstRow;
				B.Table.FirstCol = table.FirstCol;
				B.Table.LastRow = table.LastRow;
				B.Table.LastCol = table.LastCol;
				B.Table.HorizBanding = table.HorizBanding;
				B.Table.VertBanding = table.VertBanding;
				flag = true;
				goto IL_046f;
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
		int count = table.Rows.Count;
		checked
		{
			for (int i = 1; i <= count; i++)
			{
				int count2 = table.Columns.Count;
				for (int j = 1; j <= count2; j++)
				{
					try
					{
						Cell cell = B.Table.Cell(i, j);
						Cell cell2 = table.Cell(i, j);
						Font font = cell2.Shape.TextFrame.TextRange.Font;
						Font font2 = cell.Shape.TextFrame.TextRange.Font;
						font2.Bold = font.Bold;
						font2.Color.RGB = font.Color.RGB;
						font2.Italic = font.Italic;
						font2.Underline = font.Underline;
						font2.Name = font.Name;
						font2.Size = font.Size;
						_ = null;
						Microsoft.Office.Interop.PowerPoint.FillFormat fill = cell2.Shape.Fill;
						Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = cell.Shape.Fill;
						fill2.Visible = fill.Visible;
						if (fill2.Visible == MsoTriState.msoTrue)
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
							fill2.ForeColor.RGB = fill.ForeColor.RGB;
							fill2.Transparency = fill.Transparency;
						}
						fill2 = null;
						Borders borders = cell2.Borders;
						Borders borders2 = cell.Borders;
						PpBorderType[] array = new PpBorderType[4]
						{
							PpBorderType.ppBorderTop,
							PpBorderType.ppBorderBottom,
							PpBorderType.ppBorderLeft,
							PpBorderType.ppBorderRight
						};
						foreach (PpBorderType borderType in array)
						{
							borders2[borderType].Visible = borders[borderType].Visible;
							if (borders2[borderType].Visible != MsoTriState.msoTrue)
							{
								continue;
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
							borders2[borderType].ForeColor.RGB = borders[borderType].ForeColor.RGB;
							borders2[borderType].Style = borders[borderType].Style;
							borders2[borderType].DashStyle = borders[borderType].DashStyle;
							borders2[borderType].Weight = borders[borderType].Weight;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							borders2 = null;
							TextFrame textFrame = cell2.Shape.TextFrame;
							TextFrame textFrame2 = cell.Shape.TextFrame;
							textFrame2.MarginBottom = textFrame.MarginBottom;
							textFrame2.MarginLeft = textFrame.MarginLeft;
							textFrame2.MarginRight = textFrame.MarginRight;
							textFrame2.MarginTop = textFrame.MarginTop;
							textFrame2.VerticalAnchor = textFrame.VerticalAnchor;
							textFrame2.HorizontalAnchor = textFrame.HorizontalAnchor;
							textFrame2.TextRange.ParagraphFormat.Alignment = textFrame.TextRange.ParagraphFormat.Alignment;
							_ = null;
							cell2 = null;
							break;
						}
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
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_03ae;
					}
					continue;
					end_IL_03ae:
					break;
				}
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
			goto IL_046f;
		}
		IL_046f:
		checked
		{
			if (flag)
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
				int count3 = table.Rows.Count;
				for (int l = 1; l <= count3; l++)
				{
					int count4 = table.Columns.Count;
					for (int m = 1; m <= count4; m++)
					{
						try
						{
							Cell cell = B.Table.Cell(l, m);
							TextFrame textFrame = table.Cell(l, m).Shape.TextFrame;
							TextFrame textFrame3 = cell.Shape.TextFrame;
							textFrame3.MarginBottom = textFrame.MarginBottom;
							textFrame3.MarginLeft = textFrame.MarginLeft;
							textFrame3.MarginRight = textFrame.MarginRight;
							textFrame3.MarginTop = textFrame.MarginTop;
							textFrame3.VerticalAnchor = textFrame.VerticalAnchor;
							textFrame3.HorizontalAnchor = textFrame.HorizontalAnchor;
							textFrame3.TextRange.ParagraphFormat.Alignment = textFrame.TextRange.ParagraphFormat.Alignment;
							_ = null;
							_ = null;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0593;
						}
						continue;
						end_IL_0593:
						break;
					}
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
			}
			int count5 = table.Rows.Count;
			for (int n = 1; n <= count5; n++)
			{
				try
				{
					B.Table.Rows[n].Height = table.Rows[n].Height;
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
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
				int count6 = table.Columns.Count;
				for (int num = 1; num <= count6; num++)
				{
					try
					{
						B.Table.Columns[num].Width = table.Columns[num].Width;
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					B.LockAspectRatio = A.LockAspectRatio;
					table = null;
					A.Delete();
					Font font = null;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill = null;
					Borders borders = null;
					Cell cell = null;
					TextFrame textFrame = null;
					return;
				}
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
	}

	private static Slide A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Application B)
	{
		Slide result;
		try
		{
			result = B.ActivePresentation.Slides[((Slide)A.Parent).SlideIndex];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = B.ActiveWindow.Selection.SlideRange[1];
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		A.Chart.ChartData.Activate();
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, string B)
	{
		try
		{
			NewLateBinding.LateSetComplex(NewLateBinding.LateGet(A.Chart.Axes(Microsoft.Office.Core.XlAxisType.xlValue), null, AH.A(99793), new object[0], null, null, null), null, AH.A(99814), new object[1] { B }, null, null, OptimisticSet: false, RValueBase: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static string A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		string result = "";
		try
		{
			result = NewLateBinding.LateGet(NewLateBinding.LateGet(A.Chart.Axes(Microsoft.Office.Core.XlAxisType.xlValue), null, AH.A(99793), new object[0], null, null, null), null, AH.A(99814), new object[0], null, null, null).ToString();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void ViewSource(Selection sel)
	{
		int B = 0;
		try
		{
			if (sel.HasChildShapeRange)
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
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = sel.ChildShapeRange.GetEnumerator();
					while (enumerator.MoveNext())
					{
						A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_004d;
						}
						continue;
						end_IL_004d:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			else
			{
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = sel.ShapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, ref B);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_00a9;
						}
						continue;
						end_IL_00a9:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			Hyperlinks.ViewSource(sel, ref B);
			if (B > 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						Common.LogActivity(AH.A(99839));
						return;
					}
				}
			}
			Forms.WarningMessage(AH.A(95827));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref int B)
	{
		checked
		{
			if (A.Type != MsoShapeType.msoGroup)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						try
						{
							ViewSource(A);
							B++;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						try
						{
							Text.ViewSource(A);
							B++;
							return;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
							return;
						}
					}
				}
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GroupItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Shapes.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}

	public static void ViewSource(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		if (IsLinked(shp))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Link val = LinkDetails(shp);
					string text = Source.View(val);
					if (Operators.CompareString(text, val.Source, TextCompare: false) != 0)
					{
						Common.UpdateSource(shp.Tags, null, text, blnUpdateLastModified: true);
					}
					return;
				}
				}
			}
		}
		throw new Exception(AH.A(95914));
	}

	public static void BreakLinks(Selection sel)
	{
		if (!Base.ConfirmBreakLink())
		{
			return;
		}
		IEnumerator enumerator2 = default(IEnumerator);
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
			NG.A.Application.StartNewUndoEntry();
			if (!sel.HasChildShapeRange)
			{
				foreach (Microsoft.Office.Interop.PowerPoint.Shape item in sel.ShapeRange)
				{
					C(item);
				}
			}
			else
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
				try
				{
					enumerator2 = sel.ChildShapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						C((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_006e;
						}
						continue;
						end_IL_006e:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			Text.BreakLinks(sel, blnUpdateRibbon: false);
			Hyperlinks.BreakLinks(sel);
			Ribbon.LinkSelected = Ribbon.LinkSelection.No;
			clsRibbon.InvalidateLinkedItemControls();
			Common.LogActivity(AH.A(99874));
			return;
		}
	}

	private static void C(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			BreakLink(A);
			return;
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				C((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
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
				return;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	public static void BreakLink(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Common.BreakLink(shp.Tags);
	}

	public static void ConvertLegacyLink(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		if (Operators.CompareString(shp.Tags[Base.TAG_LINK_XML], string.Empty, TextCompare: false) != 0)
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
			Common.AddNewFormatLinkTag(shp.Tags, LinkDetails(shp));
			return;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, Range B)
	{
		int count = B.Rows.Count;
		int count2 = B.Columns.Count;
		Table table = A.Table;
		int count3 = table.Rows.Count;
		int count4 = table.Columns.Count;
		checked
		{
			if (count != count3)
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
				if (count > count3)
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
					int num = count - count3;
					for (int i = 1; i <= num; i++)
					{
						table.Rows.Add();
					}
				}
				else
				{
					for (int i = count3 - count; i >= 1; i += -1)
					{
						table.Rows[i].Delete();
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
				}
			}
			if (count2 != count4)
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
				if (count2 > count4)
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
					int num2 = count2 - count4;
					for (int j = 1; j <= num2; j++)
					{
						table.Columns.Add();
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
				else
				{
					for (int j = count4 - count2; j >= 1; j += -1)
					{
						table.Columns[j].Delete();
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
			}
			int count5 = table.Rows.Count;
			for (int i = 1; i <= count5; i++)
			{
				int count6 = table.Columns.Count;
				for (int j = 1; j <= count6; j++)
				{
					table.Cell(i, j).Shape.TextFrame.TextRange.Text = Conversions.ToString(NewLateBinding.LateGet(B.Cells[i, j], null, AH.A(70464), new object[0], null, null, null));
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_01dd;
					}
					continue;
					end_IL_01dd:
					break;
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				table = null;
				return;
			}
		}
	}
}
