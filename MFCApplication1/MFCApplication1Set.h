#pragma once
class CMFCApplication1Set :
    public CRecordset
{
public:
    CMFCApplication1Set(CDatabase* pds = nullptr);
    ~CMFCApplication1Set();
    virtual CString GetDefaultConnect();
    virtual CString GetDefaultSQL();
    virtual void DoFieldExchange(CFieldExchange* pFX);
    CString m_sPathName;
    CStringW* m_Title;
};

