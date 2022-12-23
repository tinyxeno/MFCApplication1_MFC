
// MFCApplication1.cpp : アプリケーションのクラス動作を定義します。
//

#include "pch.h"
#include "framework.h"
#include "afxwinappex.h"
#include "afxdialogex.h"
#include "MFCApplication1.h"
#include "MainFrm.h"

#include "MFCApplication1Set.h"
#include "MFCApplication1Doc.h"
#include "MFCApplication1View.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#if _MSC_VER < 1600
#pragma comment(lib,"version.lib")
#else
#ifdef _UNICODE
#pragma comment(lib,"Mincore.lib")
#else
#pragma comment(lib,"version.lib")
#endif
#endif


// CMFCApplication1App

BEGIN_MESSAGE_MAP(CMFCApplication1App, CWinAppEx)
	ON_COMMAND(ID_CONTEXT_HELP, &CWinAppEx::OnContextHelp)
	ON_COMMAND(ID_HELP, &CWinAppEx::OnHelp)
	ON_COMMAND(ID_HELP_FINDER, &CWinAppEx::OnHelpFinder)
	ON_COMMAND(ID_HELP_INDEX, &CWinAppEx::OnHelpIndex)
	ON_COMMAND(ID_HELP_USING, &CWinAppEx::OnHelpUsing)
	ON_COMMAND(ID_APP_ABOUT, &CMFCApplication1App::OnAppAbout)
	// 標準のファイル基本ドキュメント コマンド
	ON_COMMAND(ID_FILE_NEW, &CWinAppEx::OnFileNew)
	ON_COMMAND(ID_FILE_OPEN, &CWinAppEx::OnFileOpen)
	// 標準の印刷セットアップ コマンド
	ON_COMMAND(ID_FILE_PRINT_SETUP, &CWinAppEx::OnFilePrintSetup)
END_MESSAGE_MAP()


// CMFCApplication1App の構築

CMFCApplication1App::CMFCApplication1App() noexcept
{
	m_bHiColorIcons = TRUE;


	m_nAppLook = 0;
	// 再起動マネージャーをサポートします
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_ALL_ASPECTS;
#ifdef _MANAGED
	// アプリケーションが共通言語ランタイム サポート (/clr) を使用して作成されている場合:
	//     1) この追加設定は、再起動マネージャー サポートが正常に機能するために必要です。
	//     2) 作成するには、プロジェクトに System.Windows.Forms への参照を追加する必要があります。
	System::Windows::Forms::Application::SetUnhandledExceptionMode(System::Windows::Forms::UnhandledExceptionMode::ThrowException);
#endif

	// TODO: 下のアプリケーション ID 文字列を一意の ID 文字列で置換します。推奨される
	// 文字列の形式は CompanyName.ProductName.SubProduct.VersionInformation です
	SetAppID(_T("MFCApplication1.AppID.NoVersion"));

	// TODO: この位置に構築用コードを追加してください。
	// ここに InitInstance 中の重要な初期化処理をすべて記述してください。
	::_tsetlocale(LC_ALL, _T("japanese"));
	SetHelpMode(AFX_HELP_TYPE::afxHTMLHelp);
}

// 唯一の CMFCApplication1App オブジェクト

CMFCApplication1App theApp;


// CMFCApplication1App の初期化

BOOL CMFCApplication1App::InitInstance()
{
	// アプリケーション マニフェストが visual スタイルを有効にするために、
	// ComCtl32.dll Version 6 以降の使用を指定する場合は、
	// Windows XP に InitCommonControlsEx() が必要です。さもなければ、ウィンドウ作成はすべて失敗します。
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// アプリケーションで使用するすべてのコモン コントロール クラスを含めるには、
	// これを設定します。
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinAppEx::InitInstance();
	CString value;
	if (value.LoadString(IDS_HELP_FILE_NAME) && !value.IsEmpty())
	{
		if (nullptr != m_pszHelpFilePath)
		{
			::free((void*)m_pszHelpFilePath);
			m_pszHelpFilePath = ::_tcsdup(value);
		}
	}


	// OLE ライブラリを初期化します。
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	AfxEnableControlContainer();

	EnableTaskbarInteraction(FALSE);

	// RichEdit コントロールを使用するには AfxInitRichEdit2() が必要です
	// AfxInitRichEdit2();

	// 標準初期化
	// これらの機能を使わずに最終的な実行可能ファイルの
	// サイズを縮小したい場合は、以下から不要な初期化
	// ルーチンを削除してください。
	// 設定が格納されているレジストリ キーを変更します。
	// TODO: 会社名または組織名などの適切な文字列に
	// この文字列を変更してください。
	SetRegistryKey(_T("tinyXeono"));
	LoadStdProfileSettings(4);  // 標準の INI ファイルのオプションをロードします (MRU を含む)


	InitContextMenuManager();

	InitKeyboardManager();

	InitTooltipManager();
	CMFCToolTipInfo ttParams;
	ttParams.m_bVislManagerTheme = TRUE;
	theApp.GetTooltipManager()->SetTooltipParams(AFX_TOOLTIP_TYPE_ALL,
		RUNTIME_CLASS(CMFCToolTipCtrl), &ttParams);

	// アプリケーション用のドキュメント テンプレートを登録します。ドキュメント テンプレート
	//  はドキュメント、フレーム ウィンドウとビューを結合するために機能します。
	CSingleDocTemplate* pDocTemplate;
	pDocTemplate = new CSingleDocTemplate(
		IDR_MAINFRAME,
		RUNTIME_CLASS(CMFCApplication1Doc),
		RUNTIME_CLASS(CMainFrame),       // メイン SDI フレーム ウィンドウ
		RUNTIME_CLASS(CMFCApplication1View));
	if (!pDocTemplate)
		return FALSE;
	AddDocTemplate(pDocTemplate);


	// DDE、file open など標準のシェル コマンドのコマンド ラインを解析します。
	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);

	// DDE Execute open を使用可能にします。
	EnableShellOpen();
	RegisterShellFileTypes(TRUE);


	// コマンド ラインで指定されたディスパッチ コマンドです。アプリケーションが
	// /RegServer、/Register、/Unregserver または /Unregister で起動された場合、False を返します。
	if (!ProcessShellCommand(cmdInfo))
		return FALSE;

	// メイン ウィンドウが初期化されたので、表示と更新を行います。
	m_pMainWnd->ShowWindow(SW_SHOW);
	m_pMainWnd->UpdateWindow();
	// 接尾辞が存在する場合にのみ DragAcceptFiles を呼び出します
	//  SDI アプリケーションでは、ProcessShellCommand の直後にこの呼び出しが発生しなければなりません。
	// ドラッグ/ドロップ オープンを許可します。
	m_pMainWnd->DragAcceptFiles();
	return TRUE;
}

int CMFCApplication1App::ExitInstance()
{
	//TODO: 追加したリソースがある場合にはそれらも処理してください
	AfxOleTerm(FALSE);

	return CWinAppEx::ExitInstance();
}

// CMFCApplication1App メッセージ ハンドラー


// アプリケーションのバージョン情報に使われる CAboutDlg ダイアログ

LPCTSTR m_pAboutValueKeys[] =
{
	_T("ProductName"),
	_T("ProductVersion"),
	_T("LegalCopyright"),
	_T("CompanyName"),
	_T("FileDescription"),
};

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg() noexcept;

// ダイアログ データ
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV サポート

// 実装
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() noexcept : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	if (!pDX->m_bSaveAndValidate)
	{
		BYTE* source = nullptr;
		UINT value = 1 + MAX_PATH;
		LPTSTR result = (LPTSTR)::malloc(sizeof(TCHAR) * value);
		if (nullptr != result)
		{
			if (0 < ::GetModuleFileName(nullptr, result, (DWORD)value))
			{
				DWORD
					uDummy = 0,
					uSize = ::GetFileVersionInfoSize(result, &uDummy);
				if (0 < uSize)
				{
					source = (BYTE*)::malloc(uSize);
					if (nullptr != source)
					{
						::GetFileVersionInfo(result, 0, uSize, source);
					}
				}
			}
			::free(result);
			result = nullptr;
		}
		if (nullptr != source)
		{
			value = (UINT)(sizeof(::m_pAboutValueKeys) / sizeof(LPCTSTR));
			for (UINT index = 0; index < value; index++)
			{
				CString value = _T("\\StringFileInfo\\041104b0\\");
				value += ::m_pAboutValueKeys[index];
				UINT uDummy = 0;
				LPVOID pValue = nullptr;
				if (VerQueryValue(source, value, &pValue, (PUINT)&uDummy))
				{
					CString result = (LPCTSTR)pValue;
					CWnd* pWnd = GetDlgItem(IDC_STATIC1 + index);
					if (nullptr != pWnd)
					{
						pWnd->GetWindowText(value);
						value += result;
						pWnd->SetWindowText(value);
					}
					switch (index)
					{
					case 0:
						GetWindowText(value);
						result += value;
						SetWindowText(result);
						break;
					}
				}
			};
			::free(source);
			source = nullptr;
		}
	}
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()

// ダイアログを実行するためのアプリケーション コマンド
void CMFCApplication1App::OnAppAbout()
{
	CAboutDlg aboutDlg;
	aboutDlg.DoModal();
}

// CMFCApplication1App のカスタマイズの読み込みと保存方法

void CMFCApplication1App::PreLoadState()
{
	BOOL bNameValid;
	CString strName;
	bNameValid = strName.LoadString(IDS_EDIT_MENU);
	ASSERT(bNameValid);
	GetContextMenuManager()->AddMenu(strName, IDR_POPUP_EDIT);
}

void CMFCApplication1App::LoadCustomState()
{
}

void CMFCApplication1App::SaveCustomState()
{
}

// CMFCApplication1App メッセージ ハンドラー



