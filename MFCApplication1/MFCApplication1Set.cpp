#include "pch.h"
#include "MFCApplication1Set.h"

CMFCApplication1Set::CMFCApplication1Set(CDatabase* pdb)
	: CRecordset(pdb)
{
	m_nFields = 8;
	m_nDefaultType = CRecordset::OpenType::dynaset;
	m_Title = new CStringW[m_nFields];
}


CMFCApplication1Set::~CMFCApplication1Set()
{
	if (IsOpen())
	{
		Close();
	}
	if (nullptr != m_Title)
	{
		delete[] m_Title;
		m_Title = nullptr;
	}
}


CString CMFCApplication1Set::GetDefaultConnect()
{
	CString value;
	if (!m_sPathName.IsEmpty())
	{
		value = _T("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=");
		value += m_sPathName;
	}
	return value;
}


CString CMFCApplication1Set::GetDefaultSQL()
{
	return _T("select * from [Sheet1$]");
}


void CMFCApplication1Set::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::FieldType::outputColumn);
	for (UINT index = 0; index < m_nFields; index++)
	{
		RFX_Text(pFX, m_rgODBCFieldInfos[index].m_strName, m_Title[index]);
	};
	CRecordset::DoFieldExchange(pFX);
}
