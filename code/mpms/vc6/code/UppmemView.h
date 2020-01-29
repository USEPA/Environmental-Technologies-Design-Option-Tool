// UppmemView.h : interface of the CUppmemView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_UPPMEMVIEW_H__E1139AD5_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_)
#define AFX_UPPMEMVIEW_H__E1139AD5_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class CUppmemView : public CView
{
protected: // create from serialization only
	CUppmemView();
	DECLARE_DYNCREATE(CUppmemView)

// Attributes
public:
	CUppmemDoc* GetDocument();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CUppmemView)
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CUppmemView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CUppmemView)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // debug version in UppmemView.cpp
inline CUppmemDoc* CUppmemView::GetDocument()
   { return (CUppmemDoc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_UPPMEMVIEW_H__E1139AD5_AEBF_11D1_9A00_0020AFD5753F__INCLUDED_)
