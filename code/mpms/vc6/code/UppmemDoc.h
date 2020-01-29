// UppmemDoc.h : interface of the CUppmemDoc class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_UPPMEMDOC_H__E1139AD3_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_UPPMEMDOC_H__E1139AD3_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000


class CUppmemDoc : public CDocument
{
protected: // create from serialization only
	CUppmemDoc();
	DECLARE_DYNCREATE(CUppmemDoc)

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CUppmemDoc)
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CUppmemDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CUppmemDoc)
	afx_msg void OnModelMech();
	afx_msg void OnModelDesign();
	afx_msg void OnModelEmp();
	afx_msg void OnModelRun();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_UPPMEMDOC_H__E1139AD3_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_)
