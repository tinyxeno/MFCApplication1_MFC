
// MFCApplication1View.cpp : CMFCApplication1View クラスの実装
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
#include "MFCApplication1View.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#include "MainFrm.h"

// CMFCApplication1View

int CALLBACK SortList(LPARAM lParam1, LPARAM lParam2, LPARAM lParam3)
{
	int result = 0;
	CMainFrame* pWnd = (CMainFrame*)::AfxGetMainWnd();
	if (nullptr != pWnd)
	{
		CMFCApplication1View* pView = (CMFCApplication1View*)pWnd->GetActiveView();
		if (nullptr != pView)
		{
			result = pView->OnSortList(lParam1, lParam2, (int)lParam3);
		}
	}
	return result;
}

UINT WM_FINDREPLACE = ::RegisterWindowMessage(FINDMSGSTRING);

IMPLEMENT_DYNCREATE(CMFCApplication1View, CListView)

BEGIN_MESSAGE_MAP(CMFCApplication1View, CListView)
	// 標準印刷コマンド
	ON_COMMAND(ID_FILE_PRINT, &CListView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CListView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CMFCApplication1View::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
	ON_WM_CREATE()
	ON_NOTIFY_REFLECT(NM_SETFOCUS, &CMFCApplication1View::OnNMSetfocus)
	ON_NOTIFY_REFLECT(NM_KILLFOCUS, &CMFCApplication1View::OnNMKillfocus)
	ON_NOTIFY_REFLECT(NM_CLICK, &CMFCApplication1View::OnNMClick)
	ON_NOTIFY_REFLECT(LVN_ITEMCHANGING, &CMFCApplication1View::OnLvnItemchanging)
	ON_NOTIFY_REFLECT(LVN_KEYDOWN, &CMFCApplication1View::OnLvnKeydown)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, &CMFCApplication1View::OnNMCustomdraw)
	ON_NOTIFY_REFLECT(LVN_COLUMNCLICK, &CMFCApplication1View::OnLvnColumnclick)
	ON_NOTIFY_REFLECT(NM_DBLCLK, &CMFCApplication1View::OnNMDblclk)
	ON_NOTIFY_REFLECT(LVN_BEGINLABELEDIT, &CMFCApplication1View::OnLvnBeginlabeledit)
	ON_NOTIFY_REFLECT(LVN_ENDLABELEDIT, &CMFCApplication1View::OnLvnEndlabeledit)
	ON_WM_TIMER()
	ON_UPDATE_COMMAND_UI_RANGE(ID_VIEW_SMALLICON, ID_VIEW_DETAILS, &CMFCApplication1View::OnUpdateViewSmallicon)
	ON_COMMAND_RANGE(ID_VIEW_SMALLICON, ID_VIEW_DETAILS, &CMFCApplication1View::OnViewSmallicon)
	ON_REGISTERED_MESSAGE(WM_FINDREPLACE, &CMFCApplication1View::OnFindreplace)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_COPY, ID_EDIT_CUT, &CMFCApplication1View::OnUpdateEditCopy)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_PASTE, ID_EDIT_PASTE, &CMFCApplication1View::OnUpdateEditCopy)
	ON_UPDATE_COMMAND_UI_RANGE(ID_EDIT_SELECT_ALL, ID_EDIT_UNDO, &CMFCApplication1View::OnUpdateEditCopy)
	ON_COMMAND_RANGE(ID_EDIT_COPY, ID_EDIT_COPY, &CMFCApplication1View::OnEditCopy)
	ON_COMMAND_RANGE(ID_EDIT_PASTE, ID_EDIT_PASTE, &CMFCApplication1View::OnEditCopy)
	ON_COMMAND_RANGE(ID_EDIT_SELECT_ALL, ID_EDIT_UNDO, &CMFCApplication1View::OnEditCopy)
	ON_COMMAND(ID_EDIT_FIND, &CMFCApplication1View::OnEditFind)
	ON_COMMAND(ID_EDIT_REPEAT, &CMFCApplication1View::OnEditRepeat)
	ON_COMMAND(ID_EDIT_REPLACE, &CMFCApplication1View::OnEditReplace)
	ON_COMMAND(ID_OLE_INSERT_NEW, &CMFCApplication1View::OnOleInsertNew)
	ON_COMMAND(ID_OLE_EDIT_PROPERTIES, &CMFCApplication1View::OnOleEditProperties)
	ON_COMMAND(ID_EDIT_CLEAR, &CMFCApplication1View::OnEditClear)
END_MESSAGE_MAP()

// CMFCApplication1View コンストラクション/デストラクション

CMFCApplication1View::CMFCApplication1View() noexcept
	: m_pSet(nullptr)
	, m_iItem(0)
	, m_iSubItem(0)
	, m_direction(FALSE)
	, m_pFindReplaceDialog(nullptr)
{
	// TODO: 構築コードをここに追加します。

}

CMFCApplication1View::~CMFCApplication1View()
{
	if (nullptr != m_pFindReplaceDialog)
	{
		m_pFindReplaceDialog->SendMessage(WM_CLOSE, 0, 0);
		m_pFindReplaceDialog = nullptr;
	}
}

BOOL CMFCApplication1View::PreCreateWindow(CREATESTRUCT& cs)
{
	cs.style &= ~LVS_TYPEMASK;
	cs.style |= LVS_EDITLABELS;
	cs.style |= LVS_REPORT;
	cs.style |= LVS_SINGLESEL;
	cs.style |= LVS_SHAREIMAGELISTS;
	cs.style |= LVS_SHOWSELALWAYS;
	return CListView::PreCreateWindow(cs);
}


void CMFCApplication1View::OnDraw(CDC* /*pDC*/)
{
	CMFCApplication1Doc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	// TODO: この場所にネイティブ データ用の描画コードを追加します。
}


void CMFCApplication1View::OnInitialUpdate()
{
	CListCtrl& source = GetListCtrl();
	m_pSet = GetDocument()->m_pMFCApplication1Set;
	CMainFrame* pWnd = (CMainFrame*)::AfxGetMainWnd();
	if (nullptr != pWnd)
	{
		source.SetImageList(&pWnd->m_imgLarge, LVSIL_NORMAL);
		source.SetImageList(&pWnd->m_imgSmall, LVSIL_SMALL);
	}
	CListView::OnInitialUpdate();


	m_iItem = m_iSubItem = 0;
	BOOL result = 0 < source.GetItemCount();
	source.EnableWindow(result);
	if (result)
	{
		source.SetItemState(m_iItem,
			LVIS_SELECTED | LVIS_FOCUSED,
			LVIS_SELECTED | LVIS_FOCUSED);
		source.EnsureVisible(m_iItem, TRUE);
		source.SetFocus();
	}
}


// CMFCApplication1View の印刷


void CMFCApplication1View::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

BOOL CMFCApplication1View::OnPreparePrinting(CPrintInfo* pInfo)
{
	// 既定の印刷準備
	return DoPreparePrinting(pInfo);
}

void CMFCApplication1View::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: 印刷前の特別な初期化処理を追加してください。
}

void CMFCApplication1View::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: 印刷後の後処理を追加してください。
}

void CMFCApplication1View::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void CMFCApplication1View::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}


// CMFCApplication1View の診断

#ifdef _DEBUG
void CMFCApplication1View::AssertValid() const
{
	CListView::AssertValid();
}

void CMFCApplication1View::Dump(CDumpContext& dc) const
{
	CListView::Dump(dc);
}

CMFCApplication1Doc* CMFCApplication1View::GetDocument() const // デバッグ以外のバージョンはインラインです。
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CMFCApplication1Doc)));
	return (CMFCApplication1Doc*)m_pDocument;
}
#endif //_DEBUG


// CMFCApplication1View メッセージ ハンドラー


void CMFCApplication1View::OnPrint(CDC* pDC, CPrintInfo* pInfo)
{
	// TODO: ここに特定なコードを追加するか、もしくは基底クラスを呼び出してください。

	CListView::OnPrint(pDC, pInfo);
}


int CMFCApplication1View::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CListView::OnCreate(lpCreateStruct) == -1)
		return -1;

	CListCtrl& source = GetListCtrl();
	UINT value = source.GetExtendedStyle();
	value |= LVS_EX_FULLROWSELECT;
	value |= LVS_EX_GRIDLINES;
	value |= LVS_EX_HEADERDRAGDROP;
	value |= LVS_EX_INFOTIP;
	value |= LVS_EX_SUBITEMIMAGES;
	source.SetExtendedStyle(value);
	return 0;
}


void CMFCApplication1View::OnNMSetfocus(NMHDR* pNMHDR, LRESULT* pResult)
{
	*pResult = 0;
	CListCtrl& source = GetListCtrl();
	if (source.IsWindowVisible() && source.IsWindowEnabled() && 0 < source.GetItemCount())
	{
		source.SetFocus();
		source.Invalidate();
	}
}


void CMFCApplication1View::OnNMKillfocus(NMHDR* pNMHDR, LRESULT* pResult)
{
	*pResult = 0;
	CListCtrl& source = GetListCtrl();
	if (source.IsWindowVisible() && source.IsWindowEnabled())
	{
		source.Invalidate();
	}
}


void CMFCApplication1View::OnNMClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	*pResult = 0;
	if (nullptr != pNMItemActivate)
	{
		m_iItem = pNMItemActivate->iItem;
		m_iSubItem = pNMItemActivate->iSubItem;
		GetListCtrl().Invalidate();
	}
}


void CMFCApplication1View::OnLvnItemchanging(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	*pResult = 0;
	if (nullptr != pNMLV)
	{
		CListCtrl& source = GetListCtrl();
		switch (source.GetView())
		{
		case LVS_REPORT:
			break;
		default:
			m_iSubItem = 0;
			break;
		}
		if (m_iItem != pNMLV->iItem)
		{
			m_iItem = pNMLV->iItem;
			source.Invalidate();
		}
	}
}


void CMFCApplication1View::OnLvnKeydown(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	*pResult = 0;
	if (nullptr != pLVKeyDow)
	{
		switch (pLVKeyDow->wVKey)
		{
		case VK_LEFT:
		case VK_RIGHT:
			if (TRUE)
			{
				CListCtrl& source = GetListCtrl();
				int count = 0;
				CHeaderCtrl* index = source.GetHeaderCtrl();
				if (nullptr != index)
				{
					count = index->GetItemCount();
				}
				if (1 < count)
				{
					count--;
					LPLVCOLUMN pLvColumn = new LVCOLUMN;
					if (nullptr != pLvColumn)
					{
						pLvColumn->mask = LVCF_ORDER;
						::SecureZeroMemory(pLvColumn, sizeof(LVCOLUMN));
						if (source.GetColumn(m_iSubItem, pLvColumn))
						{
							int value = pLvColumn->iOrder;
							switch (pLVKeyDow->wVKey)
							{
							case VK_LEFT:
								if (0 < value)
								{
									value--;
								}
								break;
							case VK_RIGHT:
								if (value < count)
								{
									value++;
								}
								break;
							}
							if (pLvColumn->iOrder != value)
							{
								BOOL result = FALSE;
								for (int index = 0; false == result && index <= count; index++)
								{
									if (source.GetColumn(index, pLvColumn))
									{
										result = pLvColumn->iOrder == value;
										if (result)
										{
											m_iSubItem = index;
											source.Invalidate();
										}
									}
								};
							}
						}
						delete pLvColumn;
						pLvColumn = nullptr;
					}
				}
			}
			break;
		}
	}
}


void CMFCApplication1View::OnUpdate(CView* /*pSender*/, LPARAM /*lHint*/, CObject* /*pHint*/)
{
	CListCtrl& source = GetListCtrl();
	int count = 0;
	CHeaderCtrl* index = source.GetHeaderCtrl();
	if (nullptr != index)
	{
		count = index->GetItemCount();
	}
	BOOL result = source.DeleteAllItems();
	for (int index = count - 1; result && 0 <= index; index--)
	{
		result = source.DeleteColumn(index);
	};
	if (result && !GetDocument()->GetPathName().IsEmpty() && nullptr != m_pSet && m_pSet->IsOpen())
	{
		CString value;
		count = (int)m_pSet->m_nFields;
		for (int index = 0; result && index < count; index++)
		{
			value = m_pSet->m_rgODBCFieldInfos[index].m_strName;
			result = index == source.InsertColumn(index, value, LVCFMT_LEFT, 100);
		};
		for (m_iItem = 0; result && !m_pSet->IsEOF(); m_iItem++)
		{
			result = UpdateData(FALSE);
			if (result)
			{
				m_pSet->MoveNext();
			}
		};
	}
}


void CMFCApplication1View::DoDataExchange(CDataExchange* pDX)
{
	int count = (int)m_pSet->m_nFields;
	for (int index = 0; index < count; index++)
	{
		CString value = m_pSet->m_Title[index];
		DDX_ListViewItem(pDX, index, value);
	};
	CListView::DoDataExchange(pDX);
}


void CMFCApplication1View::OnNMCustomdraw(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	*pResult = 0;
	if (nullptr != pNMCD)
	{
		switch(pNMCD->dwDrawStage)
		{
		case CDDS_PREPAINT:
			*pResult = CDRF_NOTIFYITEMDRAW;
			break;
		case CDDS_ITEMPREPAINT:
			*pResult = CDRF_NOTIFYSUBITEMDRAW;
			break;
		case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
			*pResult = CDRF_NEWFONT;
			if (TRUE)
			{
				CListCtrl& list = GetListCtrl();
				LPNMLVCUSTOMDRAW pLVCD = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);
				UINT
					value = ::GetWindowLong(list.m_hWnd,GWL_STYLE),
					fore = ::GetSysColor(COLOR_WINDOWTEXT),
					back = ::GetSysColor(COLOR_WINDOW);
				switch (value & LVS_EDITLABELS)
				{
				case LVS_EDITLABELS:
					if (1 & pNMCD->dwItemSpec)
					{
						back = RGB(204, 255, 255);
					}
					if (m_iItem != pNMCD->dwItemSpec && m_iSubItem == pLVCD->iSubItem)
					{
						switch (LVS_SHOWSELALWAYS & value)
						{
						case LVS_SHOWSELALWAYS:
							back = ::GetSysColor(COLOR_MENU);
							break;
						}
						CWnd* pWnd = GetFocus();
						if (nullptr != pWnd && pWnd->m_hWnd == list.m_hWnd)
						{
							fore = ::GetSysColor(COLOR_HIGHLIGHTTEXT);
							back = ::GetSysColor(COLOR_HIGHLIGHT);
						}
					}
					break;
				}
				pLVCD->clrText = fore;
				pLVCD->clrTextBk = back;
			}
			break;
		}
	}
}


void CMFCApplication1View::OnLvnColumnclick(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	*pResult = 0;
	m_direction = false;
	int result = 0;
	CListCtrl& source = GetListCtrl();
	CHeaderCtrl* head = source.GetHeaderCtrl();
	if (nullptr != head)
	{
		result = head->GetItemCount();
	}
	LPLVCOLUMN pLvColumn = new LVCOLUMN;
	if (nullptr != pLvColumn)
	{
		pLvColumn->mask = LVCF_FMT;
		for (int index = 0; index < result; index++)
		{
			if (source.GetColumn(index, pLvColumn))
			{
				UINT
					mask = (HDF_SORTDOWN | HDF_SORTUP),
					value = mask;
				value &= pLvColumn->fmt;
				pLvColumn->fmt &= ~mask;
				if (index == pNMLV->iSubItem)
				{
					switch (value)
					{
					case HDF_SORTUP:
						pLvColumn->fmt |= HDF_SORTDOWN;
						m_direction = true;
						break;
					default:
						pLvColumn->fmt |= HDF_SORTUP;
						break;
					}
				}
				source.SetColumn(index, pLvColumn);
			}
		}
		delete pLvColumn;
		pLvColumn = nullptr;
	}
	source.SortItems(SortList, pNMLV->iSubItem);
}


void CMFCApplication1View::OnNMDblclk(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	*pResult = 0;
	OnOleEditProperties();
}


void CMFCApplication1View::OnLvnBeginlabeledit(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVDISPINFO* pDispInfo = reinterpret_cast<NMLVDISPINFO*>(pNMHDR);
	*pResult = 0;
	CListCtrl& source = GetListCtrl();
	CEdit* result = source.GetEditControl();
	if (nullptr != result)
	{
		result->SetWindowText(source.GetItemText(m_iItem, m_iSubItem));
		MessageBox(IDP_AFXBARRES_TEXT_IS_REQUIRED, MB_ICONINFORMATION);
	}
}


void CMFCApplication1View::OnLvnEndlabeledit(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVDISPINFO* pDispInfo = reinterpret_cast<NMLVDISPINFO*>(pNMHDR);
	*pResult = 0;
	if (nullptr != pDispInfo && pDispInfo->item.pszText)
	{
		if (_T('\0') != *pDispInfo->item.pszText)
		{
			GetListCtrl().SetItemText(m_iItem, m_iSubItem, pDispInfo->item.pszText);
			GetDocument()->SetModifiedFlag();
			MessageBox(IDS_EDIT_MENU, MB_ICONINFORMATION);
		}
		else
		{
			MessageBox(IDP_AFXBARRES_TEXT_IS_REQUIRED);
			SetTimer(ID_TIMER_REEDIT, 500, nullptr);
		}
	}
	else
	{
		MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
	}
}


BOOL CMFCApplication1View::PreTranslateMessage(MSG* pMsg)
{
	if (nullptr != pMsg)
	{
		switch (pMsg->message)
		{
		case WM_PAINT:
			if (TRUE)
			{
				CListCtrl& source = GetListCtrl();
				CEdit* result = source.GetEditControl();
				if (nullptr != result)
				{
					CRect value;
					if (source.GetSubItemRect(m_iItem, m_iSubItem, LVIR_LABEL, value))
					{
						result->MoveWindow(value);
					}
				}
			}
			break;
		}
	}
	return CListView::PreTranslateMessage(pMsg);
}


void CMFCApplication1View::OnTimer(UINT_PTR nIDEvent)
{
	switch (nIDEvent)
	{
	case ID_TIMER_REEDIT:
		KillTimer(nIDEvent);
		OnOleEditProperties();
		break;
	}
	CListView::OnTimer(nIDEvent);
}


void CMFCApplication1View::OnUpdateViewSmallicon(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		UINT value = LVS_REPORT;
		switch (pCmdUI->m_nID)
		{
		case ID_VIEW_SMALLICON:
			value = LVS_SMALLICON;
			break;
		case ID_VIEW_LARGEICON:
			value = LVS_ICON;
			break;
		case ID_VIEW_LIST:
			value = LVS_LIST;
			break;
		}
		pCmdUI->SetCheck(GetListCtrl().GetView() == value);
	}
}


void CMFCApplication1View::OnViewSmallicon(UINT nID)
{
	UINT value = LVS_REPORT;
	switch (nID)
	{
	case ID_VIEW_SMALLICON:
		value = LVS_SMALLICON;
		break;
	case ID_VIEW_LARGEICON:
		value = LVS_ICON;
		break;
	case ID_VIEW_LIST:
		value = LVS_LIST;
		break;
	}
	GetListCtrl().SetView(value);
}


void CMFCApplication1View::DDX_ListViewItem(CDataExchange* pDX, int iSubItem, CString& value)
{
	if (nullptr != pDX)
	{
		BOOL result = FALSE;
		CListCtrl& source = GetListCtrl();
		if (pDX->m_bSaveAndValidate)
		{
			value = source.GetItemText(m_iItem, iSubItem);
		}
		else
		{
			LPLVITEM pLvItem = new LVITEM;
			if (nullptr != pLvItem)
			{
				pLvItem->mask = LVIF_TEXT | LVIF_IMAGE;
				pLvItem->iItem = m_iItem;
				pLvItem->iSubItem = iSubItem;
				pLvItem->pszText = (LPTSTR)(LPCTSTR)value;
				pLvItem->iImage = 4;
				if (0 == iSubItem && source.GetItemCount() <= m_iItem)
				{
					pLvItem->mask |= LVIF_PARAM;
					pLvItem->lParam = m_iItem;
					result = m_iItem == source.InsertItem(pLvItem);
				}
				else
				{
					result = source.SetItem(pLvItem);
				}
				delete pLvItem;
				pLvItem = nullptr;
			}
		}
		if (!result)
		{
			pDX->Fail();
		}
	}
}


int CMFCApplication1View::OnSortList(LPARAM lParam1, LPARAM lParam2, int iSubItem)
{
	int result = 0;
	CString values[2];
	LVFINDINFO* pLvFindInfo = new LVFINDINFO;
	if (nullptr != pLvFindInfo)
	{
		pLvFindInfo->flags = LVFI_PARAM;
		CListCtrl& list = GetListCtrl();
		LPARAM params[] = { lParam1,lParam2 };
		for (int index = 0; index < 2; index++)
		{
			pLvFindInfo->lParam = params[index];
			int value = list.FindItem(pLvFindInfo);
			if (-1 != value)
			{
				values[index] = list.GetItemText(value, iSubItem);
			}
		};
		delete pLvFindInfo;
		pLvFindInfo = nullptr;
	}
	result = values[0].Compare(values[1]);
	if (m_direction)
	{
		result = -result;
	}
	return result;
}


LRESULT CMFCApplication1View::OnFindreplace(WPARAM wParam, LPARAM lParam)
{
	CFindReplaceDialog* pDlg = CFindReplaceDialog::GetNotifier(lParam);
	if (nullptr != pDlg)
	{
		if (pDlg->FindNext())
		{
			MessageBox(_T("次へ"), MB_ICONINFORMATION);
			OnEditRepeat();
		}
		else if (pDlg->ReplaceCurrent())
		{
			MessageBox(_T("置換して次へ"), MB_ICONINFORMATION);
			OnEditRepeat();
		}
		else if (pDlg->ReplaceAll())
		{
			MessageBox(_T("すべて置換"), MB_ICONINFORMATION);
		}
		else if (pDlg->IsTerminating())
		{
			MessageBox(_T("終了"), MB_ICONINFORMATION);
			m_pFindReplaceDialog = nullptr;
		}
	}
	return 0;
}


void CMFCApplication1View::OnUpdateEditCopy(CCmdUI* pCmdUI)
{
	if (nullptr != pCmdUI)
	{
		BOOL result = FALSE;
		CEdit* source = GetListCtrl().GetEditControl();
		if (nullptr != source)
		{
			int value = -1,
				index = -1;
			switch (pCmdUI->m_nID)
			{
			case ID_EDIT_COPY:
			case ID_EDIT_CUT:
				source->GetSel(value, index);
				result = -1 != value && -1 != index && value != index;
				break;
			case ID_EDIT_PASTE:
				result = ::IsClipboardFormatAvailable(CF_UNICODETEXT);
				break;
			case ID_EDIT_SELECT_ALL:
				result = TRUE;
				break;
			case ID_EDIT_UNDO:
				result = source->CanUndo();
				break;
			}
		}
		pCmdUI->Enable(result);
	}
}


void CMFCApplication1View::OnEditCopy(UINT nID)
{
	CEdit* source = GetListCtrl().GetEditControl();
	if (nullptr != source)
	{
		switch (nID)
		{
		case ID_EDIT_COPY:
			source->Copy();
			break;
		case ID_EDIT_CUT:
			source->Cut();
			break;
		case ID_EDIT_PASTE:
			source->Paste();
			break;
		case ID_EDIT_SELECT_ALL:
			source->SetSel(0,source->GetWindowTextLength());
			break;
		case ID_EDIT_UNDO:
			source->Undo();
			break;
		}
	}
}


void CMFCApplication1View::OnEditFind()
{
	if (nullptr == m_pFindReplaceDialog)
	{
		m_pFindReplaceDialog = new CFindReplaceDialog;
		if (nullptr != m_pFindReplaceDialog)
		{
			if (m_pFindReplaceDialog->Create(TRUE, _T(""), _T(""), FR_DOWN, this))
			{
			}
		}
	}
}


void CMFCApplication1View::OnEditRepeat()
{
	MessageBox(_T("次へ"), MB_ICONINFORMATION);
}


void CMFCApplication1View::OnEditReplace()
{
	if (nullptr == m_pFindReplaceDialog)
	{
		m_pFindReplaceDialog = new CFindReplaceDialog;
		if (nullptr != m_pFindReplaceDialog)
		{
			if (m_pFindReplaceDialog->Create(FALSE, _T(""), _T(""), FR_DOWN, this))
			{
			}
		}
	}
}


void CMFCApplication1View::OnOleInsertNew()
{
	// TODO: ここにコマンド ハンドラー コードを追加します。
}


void CMFCApplication1View::OnOleEditProperties()
{
	GetListCtrl().EditLabel(m_iItem);
}


void CMFCApplication1View::OnEditClear()
{
	// TODO: ここにコマンド ハンドラー コードを追加します。
}


UINT CMFCApplication1View::MessageBox(UINT nID, UINT style)
{
	UINT result = IDCANCEL;
	CMainFrame* pWnd = (CMainFrame*)::AfxGetMainWnd();
	if (nullptr != pWnd)
	{
		result = pWnd->MessageBox(nID, style);
	}
	return result;
}


UINT CMFCApplication1View::MessageBox(LPCTSTR caption, UINT style)
{
	UINT result = IDCANCEL;
	CMainFrame* pWnd = (CMainFrame*)::AfxGetMainWnd();
	if (nullptr != pWnd)
	{
		result = pWnd->MessageBox(caption, style);
	}
	return result;
}
