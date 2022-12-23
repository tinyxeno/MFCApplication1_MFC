
// MFCApplication1View.h : CMFCApplication1View クラスのインターフェイス
//

#pragma once


class CMFCApplication1View : public CListView
{
protected: // シリアル化からのみ作成します。
	CMFCApplication1View() noexcept;
	DECLARE_DYNCREATE(CMFCApplication1View)

// 属性
public:
	CMFCApplication1Doc* GetDocument() const;

// 操作
public:

// オーバーライド
public:
	virtual void OnDraw(CDC* pDC);  // このビューを描画するためにオーバーライドされます。
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual void OnInitialUpdate(); // 構築後に初めて呼び出されます。
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// 実装
public:
	virtual ~CMFCApplication1View();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif
	/// <summary>
	/// 行位置
	/// </summary>
	/// <returns>行位置</returns>
	int GetRow() { return m_iItem; }
	/// <summary>
	/// 列位置
	/// </summary>
	/// <returns>列位置</returns>
	int GetCol() { return m_iSubItem; }
	/// <summary>
	/// リストビューのソート処理
	/// </summary>
	/// <param name="lParam1">アイテム識別子１</param>
	/// <param name="lParam2">アイテム識別子２</param>
	/// <param name="iSubItem">列位置</param>
	/// <returns></returns>
	int OnSortList(LPARAM lParam1, LPARAM lParam2, int iSubItem);

protected:
	/// <summary>
	/// 行位置
	/// </summary>
	int m_iItem;
	/// <summary>
	/// 列位置
	/// </summary>
	int m_iSubItem;
	/// <summary>
	/// ソート方向
	/// </summary>
	BOOL m_direction;
	/// <summary>
	/// 検索/置換ダイアログ
	/// </summary>
	CFindReplaceDialog* m_pFindReplaceDialog;
	/// <summary>
	/// データベース
	/// </summary>
	CMFCApplication1Set* m_pSet;

// 生成された、メッセージ割り当て関数
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	DECLARE_MESSAGE_MAP()

	/// <summary>
	/// 印刷処理
	/// </summary>
	/// <param name="pDC"></param>
	/// <param name="pInfo"></param>
	virtual void OnPrint(CDC* pDC, CPrintInfo* pInfo);
	/// <summary>
	/// フォームをメモリーへ展開しようとした場合のイベント
	/// </summary>
	/// <param name="lpCreateStruct">ウィンドウ構築構造体</param>
	/// <returns>処理結果。成功した場合、0。失敗した場合、0以外を返す</returns>
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	/// <summary>
	/// リストビューがフォーカスを取得した場合の処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnNMSetfocus(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューがフォーカスを喪失した場合の処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnNMKillfocus(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューをクリックした場合の処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnNMClick(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューのアイテム選択を変更した場合の処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnLvnItemchanging(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューの更新処理
	/// </summary>
	/// <param name="">【未使用】</param>
	/// <param name="">【未使用】</param>
	/// <param name="">【未使用】</param>
	virtual void OnUpdate(CView* /*pSender*/, LPARAM /*lHint*/, CObject* /*pHint*/);
	/// <summary>
	/// データの読み書き処理
	/// </summary>
	/// <param name="pDX">データ交換クラス</param>
	virtual void DoDataExchange(CDataExchange* pDX);
	/// <summary>
	/// リストビュー上でキーを押下した場合の処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnLvnKeydown(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューのカスタムドロー処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnNMCustomdraw(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューのカラムクリック処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnLvnColumnclick(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューのダブルクリック処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnNMDblclk(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューの編集開始処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnLvnBeginlabeledit(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューの編集終了処理
	/// </summary>
	/// <param name="pNMHDR"></param>
	/// <param name="pResult"></param>
	afx_msg void OnLvnEndlabeledit(NMHDR* pNMHDR, LRESULT* pResult);
	/// <summary>
	/// リストビューのメッセージ処理
	/// </summary>
	/// <param name="pMsg">メッセージ構造体</param>
	/// <returns></returns>
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	/// <summary>
	/// ディレイワンショットタイマーイベント処理
	/// </summary>
	/// <param name="nIDEvent"></param>
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	/// <summary>
	/// リストビューの表示変更更新処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateViewSmallicon(CCmdUI* pCmdUI);
	/// <summary>
	/// リストビューの表示変更処理
	/// </summary>
	/// <param name="nID">コマンドID</param>
	afx_msg void OnViewSmallicon(UINT nID);
	/// <summary>
	/// リストビューの読み書き処理
	/// </summary>
	/// <param name="pDX">データ交換クラス</param>
	/// <param name="iSubItem">列位置</param>
	/// <param name="value">読み書き文字列</param>
	void DDX_ListViewItem(CDataExchange* pDX, int iSubItem, CString& value);
	/// <summary>
	/// 検索/置換ダイアログのコールバック処理
	/// </summary>
	/// <param name="wParam"></param>
	/// <param name="lParam"></param>
	/// <returns></returns>
	afx_msg LRESULT OnFindreplace(WPARAM wParam, LPARAM lParam);
	/// <summary>
	/// 「元に戻す」「切り取り」「コピー」「貼り付け」「すべて選択」の許可/禁止処理
	/// </summary>
	/// <param name="pCmdUI"></param>
	afx_msg void OnUpdateEditCopy(CCmdUI* pCmdUI);
	/// <summary>
	/// 「元に戻す」「切り取り」「コピー」「貼り付け」「すべて選択」の処理
	/// </summary>
	/// <param name="nID">コマンドID</param>
	afx_msg void OnEditCopy(UINT nID);
	/// <summary>
	/// 「検索」ダイアログの起動処理
	/// </summary>
	afx_msg void OnEditFind();
	/// <summary>
	/// 「次へ」処理
	/// </summary>
	afx_msg void OnEditRepeat();
	/// <summary>
	/// 「置換」ダイアログの起動処理
	/// </summary>
	afx_msg void OnEditReplace();
	/// <summary>
	/// リストビューの「挿入」処理
	/// </summary>
	afx_msg void OnOleInsertNew();
	/// <summary>
	/// リストビューの「編集」処理
	/// </summary>
	afx_msg void OnOleEditProperties();
	/// <summary>
	/// リストビューの「削除」処理
	/// </summary>
	afx_msg void OnEditClear();
	/// <summary>
	/// メッセージボックスの表示処理
	/// </summary>
	/// <param name="nID">表示文字列のID</param>
	/// <param name="style">メッセージボックスのスタイル</param>
	/// <returns>処理結果</returns>
	UINT MessageBox(UINT nID, UINT style = MB_ICONEXCLAMATION);
	/// <summary>
	/// メッセージボックスの表示処理
	/// </summary>
	/// <param name="nID">表示文字列</param>
	/// <param name="style">メッセージボックスのスタイル</param>
	/// <returns>処理結果</returns>
	UINT MessageBox(LPCTSTR caption, UINT style = MB_ICONEXCLAMATION);
};

#ifndef _DEBUG  // MFCApplication1View.cpp のデバッグ バージョン
inline CMFCApplication1Doc* CMFCApplication1View::GetDocument() const
   { return reinterpret_cast<CMFCApplication1Doc*>(m_pDocument); }
#endif

