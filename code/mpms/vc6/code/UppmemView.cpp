// UppmemView.cpp : implementation of the CUppmemView class
//

#include "stdafx.h"
#include "Uppmem.h"

#include "UppmemDoc.h"
#include "UppmemView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CUppmemView

IMPLEMENT_DYNCREATE(CUppmemView, CView)

BEGIN_MESSAGE_MAP(CUppmemView, CView)
	//{{AFX_MSG_MAP(CUppmemView)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, CView::OnFilePrintPreview)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CUppmemView construction/destruction

CUppmemView::CUppmemView()
{
	// TODO: add construction code here

}

CUppmemView::~CUppmemView()
{
}

BOOL CUppmemView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CUppmemView drawing

void CUppmemView::OnDraw(CDC* pDC)
{
	CUppmemDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);

	// TODO: add draw code for native data here
}

/////////////////////////////////////////////////////////////////////////////
// CUppmemView printing

BOOL CUppmemView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void CUppmemView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void CUppmemView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

/////////////////////////////////////////////////////////////////////////////
// CUppmemView diagnostics

#ifdef _DEBUG
void CUppmemView::AssertValid() const
{
	CView::AssertValid();
}

void CUppmemView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CUppmemDoc* CUppmemView::GetDocument() // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CUppmemDoc)));
	return (CUppmemDoc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CUppmemView message handlers
