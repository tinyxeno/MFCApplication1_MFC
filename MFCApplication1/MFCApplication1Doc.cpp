
// MFCApplication1Doc.cpp : CMFCApplication1Doc クラスの実装
//

#include "pch.h"
#include "framework.h"
// SHARED_HANDLERS は、プレビュー、縮小版、および検索フィルター ハンドラーを実装している ATL プロジェクトで定義でき、
// そのプロジェクトとのドキュメント コードの共有を可能にします。
#ifndef SHARED_HANDLERS
#include "MFCApplication1.h"
#endif

#include "MFCApplication1Set.h"
#include "MFCApplication1Doc.h"

#include <propkey.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#include "MainFrm.h"

// CMFCApplication1Doc

IMPLEMENT_DYNCREATE(CMFCApplication1Doc, CDocument)

BEGIN_MESSAGE_MAP(CMFCApplication1Doc, CDocument)
END_MESSAGE_MAP()


// CMFCApplication1Doc コンストラクション/デストラクション

CMFCApplication1Doc::CMFCApplication1Doc() noexcept
	: m_pMFCApplication1Set(nullptr)
{
	// TODO: この位置に 1 度だけ呼ばれる構築用のコードを追加してください。

}

CMFCApplication1Doc::~CMFCApplication1Doc()
{
}

BOOL CMFCApplication1Doc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	if (nullptr != m_pMFCApplication1Set)
	{
		delete m_pMFCApplication1Set;
		m_pMFCApplication1Set = nullptr;
	}
	return TRUE;
}




// CMFCApplication1Doc のシリアル化

void CMFCApplication1Doc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: 格納するコードをここに追加してください。
	}
	else
	{
		// TODO: 読み込むコードをここに追加してください。
	}
}

#ifdef SHARED_HANDLERS

//縮小版のサポート
void CMFCApplication1Doc::OnDrawThumbnail(CDC& dc, LPRECT lprcBounds)
{
	// このコードを変更してドキュメントのデータを描画します
	dc.FillSolidRect(lprcBounds, RGB(255, 255, 255));

	CString strText = _T("TODO: implement thumbnail drawing here");
	LOGFONT lf;

	CFont* pDefaultGUIFont = CFont::FromHandle((HFONT) GetStockObject(DEFAULT_GUI_FONT));
	pDefaultGUIFont->GetLogFont(&lf);
	lf.lfHeight = 36;

	CFont fontDraw;
	fontDraw.CreateFontIndirect(&lf);

	CFont* pOldFont = dc.SelectObject(&fontDraw);
	dc.DrawText(strText, lprcBounds, DT_CENTER | DT_WORDBREAK);
	dc.SelectObject(pOldFont);
}

// 検索ハンドラーのサポート
void CMFCApplication1Doc::InitializeSearchContent()
{
	CString strSearchContent;
	// ドキュメントのデータから検索コンテンツを設定します。
	// コンテンツの各部分は ";" で区切る必要があります

	// 例:      strSearchContent = _T("point;rectangle;circle;ole object;");
	SetSearchContent(strSearchContent);
}

void CMFCApplication1Doc::SetSearchContent(const CString& value)
{
	if (value.IsEmpty())
	{
		RemoveChunk(PKEY_Search_Contents.fmtid, PKEY_Search_Contents.pid);
	}
	else
	{
		CMFCFilterChunkValueImpl *pChunk = nullptr;
		ATLTRY(pChunk = new CMFCFilterChunkValueImpl);
		if (pChunk != nullptr)
		{
			pChunk->SetTextValue(PKEY_Search_Contents, value, CHUNK_TEXT);
			SetChunkValue(pChunk);
		}
	}
}

#endif // SHARED_HANDLERS

// CMFCApplication1Doc の診断

#ifdef _DEBUG
void CMFCApplication1Doc::AssertValid() const
{
	CDocument::AssertValid();
}

void CMFCApplication1Doc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG


// CMFCApplication1Doc コマンド


BOOL CMFCApplication1Doc::SaveModified()
{
	BOOL result = TRUE;
	if (IsModified())
	{
		result = FALSE;
		UINT index = IDCANCEL;
		CMainFrame* source = (CMainFrame*)::AfxGetMainWnd();
		if (nullptr != source)
		{
			CString value, result;
			if (result.LoadString(AFX_IDS_UNTITLED))
			{
				if (!GetTitle().IsEmpty())
				{
					result = GetTitle();
				}
			}
			::AfxFormatString1(value, AFX_IDP_ASK_TO_SAVE, result);
			index = source->MessageBox(value,
				MB_YESNOCANCEL | MB_ICONQUESTION | MB_DEFBUTTON3);
		}
		switch (index)
		{
		case IDYES:
			OnFileSave();
			result = !IsModified();
			break;
		case IDNO:
			result = TRUE;
			break;
		}
	}
	return result;
}


BOOL CMFCApplication1Doc::OnOpenDocument(LPCTSTR lpszPathName)
{
	if (!CDocument::OnOpenDocument(lpszPathName))
		return FALSE;

	// TODO: ここに特定な作成コードを追加してください。


	if (nullptr != m_pMFCApplication1Set)
	{
		delete m_pMFCApplication1Set;
		m_pMFCApplication1Set = nullptr;
	}
	BOOL result = FALSE;
	if (nullptr == m_pMFCApplication1Set)
	{
		m_pMFCApplication1Set = new CMFCApplication1Set;
		if (nullptr != m_pMFCApplication1Set)
		{
			m_pMFCApplication1Set->m_sPathName = lpszPathName;
			result = m_pMFCApplication1Set->Open();
		}
	}
	return result;
}


BOOL CMFCApplication1Doc::OnSaveDocument(LPCTSTR lpszPathName)
{
	//return CDocument::OnSaveDocument(lpszPathName);
	return TRUE;
}


void CMFCApplication1Doc::OnCloseDocument()
{
	if (nullptr != m_pMFCApplication1Set)
	{
		delete m_pMFCApplication1Set;
		m_pMFCApplication1Set = nullptr;
	}
	CDocument::OnCloseDocument();
}
