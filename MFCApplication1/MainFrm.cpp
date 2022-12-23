
// MainFrm.cpp : CMainFrame クラスの実装
//

#include "pch.h"
#include "framework.h"
#include "MFCApplication1.h"

#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#include "MFCApplication1Set.h"
#include "MFCApplication1Doc.h"
#include "MFCApplication1View.h"

// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWndEx)

const int  iMaxUserToolbars = 10;
const UINT uiFirstUserToolBarId = AFX_IDW_CONTROLBAR_FIRST + 40;
const UINT uiLastUserToolBarId = uiFirstUserToolBarId + iMaxUserToolbars - 1;

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWndEx)
	ON_WM_CREATE()
	ON_COMMAND(ID_VIEW_CUSTOMIZE, &CMainFrame::OnViewCustomize)
	ON_REGISTERED_MESSAGE(AFX_WM_CREATETOOLBAR, &CMainFrame::OnToolbarCreateNew)
	ON_COMMAND_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnApplicationLook)
	ON_UPDATE_COMMAND_UI_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnUpdateApplicationLook)
	ON_UPDATE_COMMAND_UI(ID_FILE_OPEN, &CMainFrame::OnUpdateFileOpen)
	ON_UPDATE_COMMAND_UI(ID_FILE_NEW, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_FILE_CLOSE, ID_FILE_PAGE_SETUP, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_FILE_PRINT, ID_FILE_SEND_MAIL, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_CLEAR, ID_EDIT_CLEAR_ALL, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_FIND, ID_EDIT_FIND, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_REPEAT, ID_EDIT_REPLACE, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_UNDO, ID_EDIT_REDO, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI_RANGE(ID_OLE_INSERT_NEW, ID_OLE_VERB_LAST, &CMainFrame::OnUpdateFileNew)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_ROW, &CMainFrame::OnUpdateIndicatorRow)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_COL, &CMainFrame::OnUpdateIndicatorCol)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_MODIFY, &CMainFrame::OnUpdateIndicatorModify)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_DATE, &CMainFrame::OnUpdateIndicatorDate)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_TIME, &CMainFrame::OnUpdateIndicatorTime)
	ON_WM_TIMER()
	ON_WM_CLOSE()
	ON_WM_GETMINMAXINFO()
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // ステータス ライン インジケーター
	ID_INDICATOR_ROW,
	ID_INDICATOR_COL,
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
	ID_INDICATOR_MODIFY,
	ID_INDICATOR_DATE,
	ID_INDICATOR_TIME,
};

// CMainFrame コンストラクション/デストラクション

CMainFrame::CMainFrame() noexcept
{
	// TODO: メンバー初期化コードをここに追加してください。
	theApp.m_nAppLook = theApp.GetInt(_T("ApplicationLook"), ID_VIEW_APPLOOK_VS_2008);
	::SecureZeroMemory(&m_systemTime, sizeof(SYSTEMTIME));
	UINT
		source[] = {106,103,102,101,104},
		value = sizeof(source) / sizeof(UINT);
	BOOL result = m_imgSmall.Create(16, 15, ILC_COLOR24 | ILC_MASK, value, 0);
	if (result)
	{
		result = m_imgLarge.Create(32, 32, ILC_COLOR24 | ILC_MASK, value, 0);
		if (result)
		{
			for (UINT index = 0; result && index < value; index++)
			{
				result = FALSE;
				HICON value = ::LoadIcon(nullptr, MAKEINTRESOURCE(source[index]));
				if (nullptr != value && INVALID_HANDLE_VALUE != value)
				{
					result = 0 <= m_imgSmall.Add(value);
					if (result)
					{
						result = 0 <= m_imgLarge.Add(value);
					}
				}
			};
		}
	}
}

CMainFrame::~CMainFrame()
{
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWndEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	BOOL bNameValid;

	if (!m_wndMenuBar.Create(this))
	{
		TRACE0("メニュー バーを作成できませんでした\n");
		return -1;      // 作成できない場合
	}

	m_wndMenuBar.SetPaneStyle(m_wndMenuBar.GetPaneStyle() | CBRS_SIZE_DYNAMIC | CBRS_TOOLTIPS | CBRS_FLYBY);

	// アクティブになったときメニュー バーにフォーカスを移動しない
	CMFCPopupMenu::SetForceMenuFocus(FALSE);

	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_GRIPPER | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC) ||
		!m_wndToolBar.LoadToolBar(theApp.m_bHiColorIcons ? IDR_MAINFRAME_256 : IDR_MAINFRAME))
	{
		TRACE0("ツール バーの作成に失敗しました。\n");
		return -1;      // 作成できない場合
	}

	CString strToolBarName;
	bNameValid = strToolBarName.LoadString(IDS_TOOLBAR_STANDARD);
	ASSERT(bNameValid);
	m_wndToolBar.SetWindowText(strToolBarName);

	CString strCustomize;
	bNameValid = strCustomize.LoadString(IDS_TOOLBAR_CUSTOMIZE);
	ASSERT(bNameValid);
	m_wndToolBar.EnableCustomizeButton(TRUE, ID_VIEW_CUSTOMIZE, strCustomize);

	// ユーザー定義のツール バーの操作を許可します:
	InitUserToolbars(nullptr, uiFirstUserToolBarId, uiLastUserToolBarId);

	if (!m_wndStatusBar.Create(this))
	{
		TRACE0("ステータス バーの作成に失敗しました。\n");
		return -1;      // 作成できない場合
	}
	m_wndStatusBar.SetIndicators(indicators, sizeof(indicators)/sizeof(UINT));

	// TODO: ツール バーおよびメニュー バーをドッキング可能にしない場合は、この 5 つの行を削除します
	m_wndMenuBar.EnableDocking(CBRS_ALIGN_ANY);
	m_wndToolBar.EnableDocking(CBRS_ALIGN_ANY);
	EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndMenuBar);
	DockPane(&m_wndToolBar);


	// Visual Studio 2005 スタイルのドッキング ウィンドウ動作を有効にします
	CDockingManager::SetDockingMode(DT_SMART);
	// Visual Studio 2005 スタイルのドッキング ウィンドウの自動非表示動作を有効にします
	EnableAutoHidePanes(CBRS_ALIGN_ANY);
	// 固定値に基づいてビジュアル マネージャーと visual スタイルを設定します
	OnApplicationLook(theApp.m_nAppLook);

	// ツール バーとドッキング ウィンドウ メニューの配置変更を有効にします
	EnablePaneMenu(TRUE, ID_VIEW_CUSTOMIZE, strCustomize, ID_VIEW_TOOLBAR);

	// ツール バーのクイック (Alt キーを押しながらドラッグ) カスタマイズを有効にします
	CMFCToolBar::EnableQuickCustomization();

	if (CMFCToolBar::GetUserImages() == nullptr)
	{
		// ユーザー定義のツール バー イメージを読み込みます
		if (m_UserImages.Load(_T(".\\UserImages.bmp")))
		{
			CMFCToolBar::SetUserImages(&m_UserImages);
		}
	}

	// メニューのパーソナル化 (最近使用されたコマンド) を有効にします
	// TODO: ユーザー固有の基本コマンドを定義し、各メニューをクリックしたときに基本コマンドが 1 つ以上表示されるようにします。
	CList<UINT, UINT> lstBasicCommands;

	lstBasicCommands.AddTail(ID_FILE_NEW);
	lstBasicCommands.AddTail(ID_FILE_OPEN);
	lstBasicCommands.AddTail(ID_FILE_SAVE);
	lstBasicCommands.AddTail(ID_FILE_PRINT);
	lstBasicCommands.AddTail(ID_APP_EXIT);
	lstBasicCommands.AddTail(ID_EDIT_CUT);
	lstBasicCommands.AddTail(ID_EDIT_PASTE);
	lstBasicCommands.AddTail(ID_EDIT_UNDO);
	lstBasicCommands.AddTail(ID_APP_ABOUT);
	lstBasicCommands.AddTail(ID_VIEW_STATUS_BAR);
	lstBasicCommands.AddTail(ID_VIEW_TOOLBAR);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2003);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_VS_2005);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_BLUE);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_SILVER);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_BLACK);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_OFF_2007_AQUA);
	lstBasicCommands.AddTail(ID_VIEW_APPLOOK_WINDOWS_7);

	CMFCToolBar::SetBasicCommands(lstBasicCommands);

	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWndEx::PreCreateWindow(cs) )
		return FALSE;
	// TODO: この位置で CREATESTRUCT cs を修正して Window クラスまたはスタイルを
	//  修正してください。

	return TRUE;
}

// CMainFrame の診断

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWndEx::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWndEx::Dump(dc);
}
#endif //_DEBUG


// CMainFrame メッセージ ハンドラー

void CMainFrame::OnViewCustomize()
{
	CMFCToolBarsCustomizeDialog* pDlgCust = new CMFCToolBarsCustomizeDialog(this, TRUE /* メニューをスキャンします*/);
	pDlgCust->EnableUserDefinedToolbars();
	pDlgCust->Create();
}

LRESULT CMainFrame::OnToolbarCreateNew(WPARAM wp,LPARAM lp)
{
	LRESULT lres = CFrameWndEx::OnToolbarCreateNew(wp,lp);
	if (lres == 0)
	{
		return 0;
	}

	CMFCToolBar* pUserToolbar = (CMFCToolBar*)lres;
	ASSERT_VALID(pUserToolbar);

	BOOL bNameValid;
	CString strCustomize;
	bNameValid = strCustomize.LoadString(IDS_TOOLBAR_CUSTOMIZE);
	ASSERT(bNameValid);

	pUserToolbar->EnableCustomizeButton(TRUE, ID_VIEW_CUSTOMIZE, strCustomize);
	return lres;
}

void CMainFrame::OnApplicationLook(UINT id)
{
	CWaitCursor wait;

	theApp.m_nAppLook = id;

	switch (theApp.m_nAppLook)
	{
	case ID_VIEW_APPLOOK_WIN_2000:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManager));
		break;

	case ID_VIEW_APPLOOK_OFF_XP:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOfficeXP));
		break;

	case ID_VIEW_APPLOOK_WIN_XP:
		CMFCVisualManagerWindows::m_b3DTabsXPTheme = TRUE;
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));
		break;

	case ID_VIEW_APPLOOK_OFF_2003:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2003));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	case ID_VIEW_APPLOOK_VS_2005:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2005));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	case ID_VIEW_APPLOOK_VS_2008:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2008));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	case ID_VIEW_APPLOOK_WINDOWS_7:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows7));
		CDockingManager::SetDockingMode(DT_SMART);
		break;

	default:
		switch (theApp.m_nAppLook)
		{
		case ID_VIEW_APPLOOK_OFF_2007_BLUE:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_LunaBlue);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_BLACK:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_ObsidianBlack);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_SILVER:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Silver);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_AQUA:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Aqua);
			break;
		}

		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2007));
		CDockingManager::SetDockingMode(DT_SMART);
	}

	RedrawWindow(nullptr, nullptr, RDW_ALLCHILDREN | RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME | RDW_ERASE);

	theApp.WriteInt(_T("ApplicationLook"), theApp.m_nAppLook);
}

void CMainFrame::OnUpdateApplicationLook(CCmdUI* pCmdUI)
{
	pCmdUI->SetRadio(theApp.m_nAppLook == pCmdUI->m_nID);
}


BOOL CMainFrame::LoadFrame(UINT nIDResource, DWORD dwDefaultStyle, CWnd* pParentWnd, CCreateContext* pContext)
{
	// 基底クラスが実際の動作を行います

	if (!CFrameWndEx::LoadFrame(nIDResource, dwDefaultStyle, pParentWnd, pContext))
	{
		return FALSE;
	}


	// すべてのユーザー定義ツール バーのボタンのカスタマイズを有効にします
	BOOL bNameValid;
	CString strCustomize;
	bNameValid = strCustomize.LoadString(IDS_TOOLBAR_CUSTOMIZE);
	ASSERT(bNameValid);

	for (int i = 0; i < iMaxUserToolbars; i ++)
	{
		CMFCToolBar* pUserToolbar = GetUserToolBarByIndex(i);
		if (pUserToolbar != nullptr)
		{
			pUserToolbar->EnableCustomizeButton(TRUE, ID_VIEW_CUSTOMIZE, strCustomize);
		}
	}
	// ツールバーのボタンへタイトル文字列を設定する
	int value = m_wndToolBar.GetCount();
	for (int index = 0; index < value; index++)
	{
		CString value;
		UINT nID = 0, nStyle = 0;
		int iImage = 0;
		m_wndToolBar.GetButtonInfo(index, nID, nStyle, iImage);
		switch (nID)
		{
		case ID_SEPARATOR:
			break;
		default:
			if (value.LoadString(nID))
			{
				BOOL result = AfxExtractSubString(value, value, 2);
				if (!result)
				{
					result = AfxExtractSubString(value, value, 1);
				}
				if (result)
				{
					m_wndToolBar.SetButtonText(index, value);
				}
			}
			break;
		}
	};
	// ステータスバーのペインの幅を最低に設定する
	int index = m_wndStatusBar.CommandToIndex(ID_SEPARATOR);
	if (IDC_STATIC != index)
	{
		m_wndStatusBar.SetPaneWidth(index, 1);
	}
	// 起動時のメインフォームの大きさを設定
	SetWindowPos(nullptr, 0, 0, 838, 340, SWP_NOMOVE | SWP_NOZORDER);
	// 準備完了を表示
	MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
	// インターバルタイマーを起動
	SetTimer(ID_TIMER_INTERVAL, 128, nullptr);
	return TRUE;
}



void CMainFrame::OnUpdateFileOpen(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = pDoc->GetPathName().IsEmpty();
		}
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateFileNew(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = !pDoc->GetPathName().IsEmpty();
		}
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateIndicatorRow(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		int index = 0;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = !pDoc->GetPathName().IsEmpty();
			if (result)
			{
				CMFCApplication1View* pView =
					(CMFCApplication1View*)GetActiveView();
				if (nullptr != pView)
				{
					index = 1 + pView->GetRow();
				}
			}
		}
		CString value;
		value.Format(ID_INDICATOR_ROW, index);
		pCmdUI->SetText(value);
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateIndicatorCol(CCmdUI* pCmdUI)
{
	BOOL result = FALSE;
	int index = 0;
	CDocument* pDoc = GetActiveDocument();
	if (nullptr != pDoc)
	{
		result = !pDoc->GetPathName().IsEmpty();
		if (result)
		{
			CMFCApplication1View* pView =
				(CMFCApplication1View*)GetActiveView();
			if (nullptr != pView)
			{
				index = 1 + pView->GetCol();
			}
		}
	}
	CString value;
	value.Format(ID_INDICATOR_COL, index);
	pCmdUI->SetText(value);
	pCmdUI->Enable(result);
}


void CMainFrame::OnUpdateIndicatorModify(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			result = pDoc->IsModified();
		}
		pCmdUI->Enable(result);
	}
}


void CMainFrame::OnUpdateIndicatorDate(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		CString value;
		value.Format(_T("%04d/%02d/%02d"),
			m_systemTime.wYear,
			m_systemTime.wMonth,
			m_systemTime.wDay);
		pCmdUI->SetText(value);
		pCmdUI->Enable(TRUE);
	}
}


void CMainFrame::OnUpdateIndicatorTime(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		CString value;
		value.Format(
			m_systemTime.wMilliseconds < 500 ?
			_T("%02d:%02d") : _T("%02d %02d"),
			m_systemTime.wHour, m_systemTime.wMinute);
		pCmdUI->SetText(value);
		pCmdUI->Enable(TRUE);
	}
}


void CMainFrame::OnTimer(UINT_PTR nIDEvent)
{
	switch (nIDEvent)
	{
	case ID_TIMER_INTERVAL:
		::GetLocalTime(&m_systemTime);
		break;
	}
	CFrameWndEx::OnTimer(nIDEvent);
}


void CMainFrame::OnClose()
{
	BOOL result = TRUE;
	if (m_wndToolBar.IsWindowVisible())
	{
		CDocument* pDoc = GetActiveDocument();
		if (nullptr != pDoc)
		{
			if (pDoc->IsModified())
			{
				result = pDoc->SaveModified();
				if (result)
				{
					pDoc->SetModifiedFlag(FALSE);
				}
			}
			else
			{
				result = IDYES ==
					MessageBox(AFX_IDP_ASK_TO_EXIT,
						MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2);
			}
			if (result)
			{
				KillTimer(ID_TIMER_INTERVAL);
			}
		}
	}
	if (result)
	{
		CFrameWndEx::OnClose();
	}
}


UINT CMainFrame::MessageBox(UINT nID, UINT style)
{
	BOOL result = IDOK;
	CString value;
	if (value.LoadString(nID))
	{
		result = MessageBox(value, style);
	}
	return result;
}


UINT CMainFrame::MessageBox(LPCTSTR caption, UINT style)
{
	UINT index = MB_ICONEXCLAMATION | MB_ICONINFORMATION;
	index &= style;
	index >>= 4;
	{
		HICON value = nullptr;
		int result = m_wndStatusBar.CommandToIndex(ID_SEPARATOR);
		switch (result)
		{
		case IDC_STATIC:
			break;
		default:
			if ((int)index < m_imgSmall.GetImageCount())
			{
				value = m_imgSmall.ExtractIcon(index);
			}
			if (nullptr == value)
			{
				value = m_imgSmall.ExtractIcon(4);
			}
			if (nullptr != value)
			{
				m_wndStatusBar.SetPaneIcon(result, value);
			}
			break;
		}
	}
	CString value;
	if (value.LoadString(IDS_EDIT_TITLE))
	{
		::AfxExtractSubString(value, value, index);
	}
	SendMessage(WM_SETMESSAGESTRING, 0, (LPARAM)caption);
	BOOL result = IDOK;
	switch (style)
	{
	case MB_ICONINFORMATION:
		break;
	default:
		result = CWnd::MessageBox(caption, value, style);
		break;
	}
	switch (result)
	{
	case IDCANCEL:
	case IDNO:
		MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
		break;
	case IDABORT:
		MessageBox(IDS_EDIT_ABORT, MB_ICONINFORMATION);
		break;
	}
	return result;
}

void CMainFrame::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	lpMMI->ptMinTrackSize.x = 570;
	lpMMI->ptMinTrackSize.y = 169;
	CFrameWnd::OnGetMinMaxInfo(lpMMI);
}

