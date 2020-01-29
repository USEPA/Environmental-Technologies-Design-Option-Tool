//**********************************************************************************************************
//        
//        MOlecular Structure DissAssembly Program (MOSDAP) 2.0
//        (Header file for:
//							Graph Match Structure
//							Hash Comparison Classes
//							Subfragment Location Classes
//		  )	
//        5/98 
//**********************************************************************************************************
//
//        Copyright (c) by John W. Raymond, Jr., 1998
//        All rights reserved.
//        
//**********************************************************************************************************

#if !defined(H_SDA_2)
#define H_SDA_2

#include "SDA_1.H"

//Qeue node class used in qeue for Figueras (1996) ring perception routine.
class QueNode {
public:
	unsigned _int32 intSource;
	unsigned _int32 intAtom_Index;
	QueNode *ptrNextQue;

	void Add_Que(unsigned _int32 intSource_Node,unsigned _int32 intIndex_Node);
	QueNode();
	~QueNode();
};

//Class used as the header for the hash table
class HashHeader {
public:
     unsigned _int32 intChainLength;
     unsigned _int32 *ptrSepChain;
     HashHeader();
     ~HashHeader();
};

//Hash Table structure class
class HashStructure {
public:
	unsigned _int8 intHash_PN;				//Prime number used in hashing
	unsigned _int8 intHash_Flag;			//Flag denoting whether hashing should be used in atom matching (0~no,1~yes).
	HashHeader *ptrHash_Bucket;
	HashStructure();
	~HashStructure();

	void Reset_HT();
	void Hash_Graph(Atom *ptrStructure,unsigned _int32 intNum_Atoms);
};

//Match element (node) class used in creating the linked list matching structure. 
class MatchElement {
public:

	 void Add_MN(unsigned _int32 intQuery_Loc);
	 unsigned _int8 Delete_MN(MatchElement *&ptrOld_Match);
	
	 unsigned _int32 intMatch;
     MatchElement *ptrNextMatch;
     MatchElement();
     ~MatchElement();
};

class MatchStructure {
public:
	
	MatchElement **ptrMS_Spine;		//Master match element array
	MatchElement ***ptrRefine_Rack;	//Diagonal holding structure for refinement procedure
	MatchElement ***ptrRefine_Top;	//Stores last match element address for filling rack
	MatchElement **ptrCol_Loc;		//Stores location of selected query(alpha) elements
	SubFragLoc *ptrAlpha_Loc;

	MatchStructure(unsigned _int32 intNum_Beta_Atoms,unsigned _int32 intNum_Alpha_Atoms,unsigned _int8 intSearch_Type);
	~MatchStructure();

	void Clear_MS(unsigned _int32 intNum_Beta_Atoms);
	void New_MN(MatchElement *&ptrTemp3,unsigned _int32 intDepth,unsigned _int32 intQuery_Loc);
	unsigned _int8 Rack_MN(unsigned _int32 intBeta_Row,unsigned _int32 intBeta_Depth,MatchElement *&ptrOld_Match,MatchElement *&ptrTempMatch);
	void UnRack_MN(unsigned _int32 intBeta_Depth,unsigned _int32 intNum_Beta_Atoms);
	void UnRack_All(unsigned _int32 intNum_Rack_Atoms);
	void Initialize_Arrays(unsigned _int32 intNum_Beta_Atoms);
	unsigned _int8 Truncate_Spine(ChemSeqID *ptrAlpha,ChemSeqID *ptrBeta);
	
};

#endif    //H_SDA_2
