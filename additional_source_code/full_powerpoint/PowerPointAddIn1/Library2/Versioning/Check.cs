using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Links;

namespace PowerPointAddIn1.Library2.Versioning;

public sealed class Check
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<SlideItem, bool> A;

		public static Func<ShapeItem, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(SlideItem A)
		{
			return ((ContentItem)A).IsOutdated;
		}

		[SpecialName]
		internal bool A(ShapeItem A)
		{
			return ((ContentItem)A).IsOutdated;
		}
	}

	[CompilerGenerated]
	internal sealed class JD
	{
		public Microsoft.Office.Interop.PowerPoint.Presentation A;

		public Action A;

		[SpecialName]
		internal void A()
		{
			try
			{
				Dispatcher a = Check.m_A;
				Action callback;
				if (this.A != null)
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
					callback = this.A;
				}
				else
				{
					callback = (this.A = [SpecialName] () =>
					{
						Check.A(this.A, B: false);
						List<SlideItem> librarySlides = LibrarySlides;
						Func<SlideItem, bool> predicate;
						if (_Closure_0024__.A == null)
						{
							predicate = (_Closure_0024__.A = _Closure_0024__.A.A);
						}
						else
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
							predicate = _Closure_0024__.A;
						}
						if (librarySlides.Where(predicate).Count() <= 0)
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
							if (LibraryShapes.Where(_Closure_0024__.A.A).Count() <= 0)
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
								break;
							}
						}
						if (UIFormsExtensions.AskYesNo((Window)null, AH.A(170624), true, true))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									Check.A();
									return;
								}
							}
						}
					});
				}
				a.Invoke(callback);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}

		[SpecialName]
		internal void B()
		{
			Check.A(this.A, B: false);
			List<SlideItem> librarySlides = LibrarySlides;
			Func<SlideItem, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = _Closure_0024__.A.A);
			}
			else
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
				predicate = _Closure_0024__.A;
			}
			if (librarySlides.Where(predicate).Count() <= 0)
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
				if (LibraryShapes.Where(_Closure_0024__.A.A).Count() <= 0)
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
					break;
				}
			}
			if (!UIFormsExtensions.AskYesNo((Window)null, AH.A(170624), true, true))
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
				Check.A();
				return;
			}
		}
	}

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static List<SlideItem> m_A;

	[CompilerGenerated]
	private static List<ShapeItem> m_A;

	private static Dispatcher m_A;

	internal static bool CheckOutdatedLibraryContent
	{
		[CompilerGenerated]
		get
		{
			return Check.m_A;
		}
		[CompilerGenerated]
		set
		{
			Check.m_A = value;
		}
	}

	internal static List<SlideItem> LibrarySlides
	{
		[CompilerGenerated]
		get
		{
			return Check.m_A;
		}
		[CompilerGenerated]
		set
		{
			Check.m_A = value;
		}
	}

	internal static List<ShapeItem> LibraryShapes
	{
		[CompilerGenerated]
		get
		{
			return Check.m_A;
		}
		[CompilerGenerated]
		set
		{
			Check.m_A = value;
		}
	}

	internal static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		if (!CheckOutdatedLibraryContent || A.Windows.Count <= 0)
		{
			return;
		}
		Action A2 = default(Action);
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
			Check.m_A = Dispatcher.CurrentDispatcher;
			Thread thread = new Thread([SpecialName] () =>
			{
				try
				{
					Dispatcher a = Check.m_A;
					Action callback;
					if (A2 != null)
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
						callback = A2;
					}
					else
					{
						callback = (A2 = [SpecialName] () =>
						{
							Check.A(A, B: false);
							List<SlideItem> librarySlides = LibrarySlides;
							Func<SlideItem, bool> predicate;
							if (_Closure_0024__.A == null)
							{
								predicate = (_Closure_0024__.A = _Closure_0024__.A.A);
							}
							else
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
								predicate = _Closure_0024__.A;
							}
							if (librarySlides.Where(predicate).Count() <= 0)
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
								if (LibraryShapes.Where(_Closure_0024__.A.A).Count() <= 0)
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
									break;
								}
							}
							if (UIFormsExtensions.AskYesNo((Window)null, AH.A(170624), true, true))
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										Check.A();
										return;
									}
								}
							}
						});
					}
					a.Invoke(callback);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			});
			thread.IsBackground = true;
			thread.SetApartmentState(ApartmentState.STA);
			thread.Start();
			return;
		}
	}

	private static void A()
	{
		try
		{
			Pane.A();
			Pane.B();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	internal static void B()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		if (IG.A(application.Presentations) > 0)
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
			A(application.ActivePresentation, B: true);
			if (LibrarySlides.Count <= 0)
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
				if (LibraryShapes.Count <= 0)
				{
					Forms.InfoMessage(AH.A(59101));
					goto IL_008b;
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
			Pane.A();
			goto IL_008b;
		}
		Forms.WarningMessage(AH.A(59235));
		goto IL_00ad;
		IL_00ad:
		application = null;
		return;
		IL_008b:
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(59194));
		goto IL_00ad;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, bool B)
	{
		List<SlideItem> B2 = new List<SlideItem>();
		List<ShapeItem> B3 = new List<ShapeItem>();
		try
		{
			IEnumerator enumerator = A.Slides.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					Check.A(slide, ref B2, B);
					foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
					{
						Check.A(shape, ref B3);
					}
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			if (B)
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
				Forms.ErrorMessage(ex2.Message);
			}
			ProjectData.ClearProjectError();
		}
		finally
		{
			LibrarySlides = B2;
			LibraryShapes = B3;
		}
		B2 = null;
		B3 = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref List<ShapeItem> B)
	{
		if (A.Type == MsoShapeType.msoGroup)
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
			if (!Tagging.A(A))
			{
				{
					IEnumerator enumerator = A.GroupItems.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Check.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B);
						}
						while (true)
						{
							switch (6)
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
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
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
		}
		Check.B(A, ref B);
	}

	private static void A(Slide A, ref List<SlideItem> B, bool C)
	{
		//IL_0114: Unknown result type (might be due to invalid IL or missing references)
		//IL_0119: Unknown result type (might be due to invalid IL or missing references)
		//IL_011b: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0135: Unknown result type (might be due to invalid IL or missing references)
		//IL_013a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0141: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0192: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		ContentInfo? val = Tagging.A(A);
		if (val.HasValue)
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
					using (List<SlideItem>.Enumerator enumerator = B.GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							SlideItem current = enumerator.Current;
							if (Operators.CompareString(((ContentItem)current).ContentInfo.ContentId, val.Value.ContentId, TextCompare: false) == 0)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										current.A(A);
										return;
									}
								}
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0082;
							}
							continue;
							end_IL_0082:
							break;
						}
					}
					ContentInfo value = val.Value;
					string libraryPath = Content.GetLibraryPath(value);
					if (Content.PathsAreDifferent(libraryPath, value))
					{
						Tagging.A(A.Tags, libraryPath);
					}
					ManifestInfo? manifestInfo = Content.GetManifestInfo(libraryPath, value);
					if (manifestInfo.HasValue)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								B.Add(new SlideItem(A, value, manifestInfo.Value));
								return;
							}
						}
					}
					return;
				}
				}
			}
		}
		if (!Check.A(A))
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
			Link val2 = Check.A(A);
			using (List<SlideItem>.Enumerator enumerator2 = B.GetEnumerator())
			{
				while (enumerator2.MoveNext())
				{
					SlideItem current2 = enumerator2.Current;
					if (Operators.CompareString(((ContentItem)current2).ContentInfo.ContentId, val2.Name, TextCompare: false) != 0)
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
						current2.A(A);
						return;
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0174;
					}
					continue;
					end_IL_0174:
					break;
				}
			}
			try
			{
				val2.Source = Check.A(A, val2);
				ContentInfo val3 = Check.A(val2);
				ManifestInfo? manifestInfo2 = Content.GetManifestInfo(val2.Source, val3);
				if (manifestInfo2.HasValue)
				{
					SlideItem slideItem = new SlideItem(A, val3, manifestInfo2.Value);
					((ContentItem)slideItem).IsOutdated = Updates.IsUpdateAvailable(val2);
					((ContentItem)slideItem).IsLegacySlideLink = true;
					SlideItem item = slideItem;
					B.Add(item);
					item = null;
				}
				return;
			}
			catch (NotSupportedException ex)
			{
				ProjectData.SetProjectError(ex);
				NotSupportedException ex2 = ex;
				if (C)
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
					Forms.ErrorMessage(AH.A(59278) + ex2.Message);
				}
				ProjectData.ClearProjectError();
				return;
			}
			catch (IOException ex3)
			{
				ProjectData.SetProjectError(ex3);
				IOException ex4 = ex3;
				if (C)
				{
					Forms.ErrorMessage(AH.A(59278) + ex4.Message);
				}
				ProjectData.ClearProjectError();
				return;
			}
			catch (UnauthorizedAccessException ex5)
			{
				ProjectData.SetProjectError(ex5);
				UnauthorizedAccessException ex6 = ex5;
				ProjectData.ClearProjectError();
				return;
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				clsReporting.LogException(ex8);
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	private static ContentInfo A(Link A)
	{
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_0085: Unknown result type (might be due to invalid IL or missing references)
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f2: Unknown result type (might be due to invalid IL or missing references)
		string id = default(string);
		Permission permission = default(Permission);
		using (List<LibraryItem>.Enumerator enumerator = Base.LibraryCollection.GetEnumerator())
		{
			while (true)
			{
				if (enumerator.MoveNext())
				{
					LibraryItem current = enumerator.Current;
					if (A.Source.StartsWith(current.Location))
					{
						id = current.Id;
						permission = current.Permission;
						break;
					}
					continue;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					break;
				}
				break;
			}
		}
		return new ContentInfo
		{
			ContentId = A.Name,
			ContentPath = A.Source,
			ContentType = (ContentType)1,
			ModifiedAt = A.LastUpdate,
			ModifiedBy = A.LastUser,
			KeepSourceFormatting = A.KeepSourceFormatting,
			LibraryId = id,
			Title = AH.A(59401),
			CurrentVersion = 1,
			IgnoredVersion = 0,
			Permission = permission
		};
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A, ref List<ShapeItem> B)
	{
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_0073: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		ContentInfo? val = Tagging.A(A);
		if (!val.HasValue)
		{
			return;
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
			ContentInfo value = val.Value;
			string libraryPath = Content.GetLibraryPath(value);
			if (Content.PathsAreDifferent(libraryPath, value))
			{
				Tagging.A(A.Tags, libraryPath);
			}
			ManifestInfo? manifestInfo = Content.GetManifestInfo(libraryPath, value);
			if (!manifestInfo.HasValue)
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
				B.Add(new ShapeItem(A, value, manifestInfo.Value));
				return;
			}
		}
	}

	internal static bool A(Slide A)
	{
		return Common.IsLinked(A.Tags);
	}

	internal static Link A(Slide A)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0246: Unknown result type (might be due to invalid IL or missing references)
		//IL_0248: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_032f: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Link val = default(Link);
		string text = default(string);
		XmlDocument xmlDocument = default(XmlDocument);
		IEnumerator enumerator = default(IEnumerator);
		XmlNode xmlNode = default(XmlNode);
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
				case 734:
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
							goto IL_009e;
						case 9:
							goto IL_00a0;
						case 10:
							goto IL_00b3;
						case 11:
							goto IL_00c6;
						case 12:
							goto IL_00d9;
						case 13:
							goto IL_00ec;
						case 14:
							goto IL_00f8;
						case 15:
							goto IL_010b;
						case 16:
							goto IL_011e;
						case 17:
						case 18:
							goto IL_0131;
						case 19:
							goto IL_014c;
						case 20:
							goto IL_016e;
						case 22:
							goto IL_0176;
						case 23:
							goto IL_0179;
						case 24:
							goto IL_019a;
						case 25:
							goto IL_01bf;
						case 26:
							goto IL_01e4;
						case 27:
							goto IL_01f3;
						case 28:
							goto IL_01ff;
						case 29:
							goto IL_0220;
						case 21:
						case 30:
						case 31:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 32:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0176:
					num2 = 22;
					goto IL_0179;
					IL_0007:
					num2 = 2;
					val = default(Link);
					goto IL_0011;
					IL_0011:
					num2 = 3;
					text = A.Tags[Base.TAG_LINK_XML];
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
					goto IL_0176;
					IL_0179:
					num2 = 23;
					val.Source = A.Tags[Base.TAG_LINK_SOURCE].ToString();
					goto IL_019a;
					IL_019a:
					num2 = 24;
					val.SourceModified = A.Tags[Base.TAG_LINK_SOURCE_LAST_MOD].ToString();
					goto IL_01bf;
					IL_01bf:
					num2 = 25;
					val.Name = A.Tags[Base.TAG_LINK_NAME].ToString();
					goto IL_01e4;
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
					enumerator = xmlDocument.DocumentElement.SelectNodes(AH.A(59442)).GetEnumerator();
					goto IL_0134;
					IL_0134:
					if (enumerator.MoveNext())
					{
						xmlNode = (XmlNode)enumerator.Current;
						goto IL_009e;
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
					goto IL_014c;
					IL_01e4:
					num2 = 26;
					val.ParentId = "";
					goto IL_01f3;
					IL_014c:
					num2 = 19;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_016e;
					IL_01ff:
					num2 = 28;
					val.LastUpdate = A.Tags[Base.TAG_LINK_TIME].ToString();
					goto IL_0220;
					IL_0220:
					num2 = 29;
					val.LastUser = A.Tags[Base.TAG_LINK_USER].ToString();
					break;
					IL_016e:
					xmlDocument = null;
					break;
					IL_01f3:
					num2 = 27;
					val.Type = (ImportType)10;
					goto IL_01ff;
					IL_009e:
					num2 = 8;
					goto IL_00a0;
					IL_00a0:
					num2 = 9;
					val.Source = Common.GetLinkSource(xmlNode);
					goto IL_00b3;
					IL_00b3:
					num2 = 10;
					val.SourceModified = Common.GetLinkSourceModified(xmlNode);
					goto IL_00c6;
					IL_00c6:
					num2 = 11;
					val.Name = Common.GetLinkId(xmlNode);
					goto IL_00d9;
					IL_00d9:
					num2 = 12;
					val.ParentId = Common.GetParentId(xmlNode);
					goto IL_00ec;
					IL_00ec:
					num2 = 13;
					val.Type = (ImportType)10;
					goto IL_00f8;
					IL_00f8:
					num2 = 14;
					val.LastUpdate = Common.GetLinkTime(xmlNode);
					goto IL_010b;
					IL_010b:
					num2 = 15;
					val.LastUser = Common.GetLinkUser(xmlNode);
					goto IL_011e;
					IL_011e:
					num2 = 16;
					val.KeepSourceFormatting = Check.A(xmlNode);
					goto IL_0131;
					IL_0131:
					num2 = 18;
					goto IL_0134;
					end_IL_0000_2:
					break;
				}
				num2 = 31;
				result = val;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 734;
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
				switch (6)
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

	private static bool A(XmlNode A)
	{
		bool result;
		try
		{
			result = Conversions.ToBoolean(Common.GetLinkKeepSourceFormatting(A));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static string A(Slide A, Link B)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		string text = Content.LibrarySourcePathFromContentId(B.Source, B.Name);
		if (Operators.CompareString(text, B.Source, TextCompare: false) != 0)
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
			Common.UpdateSource(A.Tags, null, text, blnUpdateLastModified: true);
		}
		return text;
	}
}
