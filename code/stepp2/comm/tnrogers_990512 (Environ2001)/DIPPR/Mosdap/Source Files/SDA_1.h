//**********************************************************************************************************
//        
//        MOlecular Structure DissAssembly Program (MOSDAP) 2.0
//        (Feature Searching)
//        5/98 
//**********************************************************************************************************
//
//        Copyright (c) by John W. Raymond, Jr., 1998
//        All rights reserved.
//        
//**********************************************************************************************************

#if !defined(H_SDA_1)
#define H_SDA_1

#include <math.h>
#include <iostream.h>
#include <fstream.h>
#include <string.h>

//Class used to store pointer addresses of detection(s) for subfragment(s) located in query molecule
class SubFragLoc{
public:
     unsigned _int32 *ptrQuery_Loc;
	 unsigned _int32 intArray_Length;
     SubFragLoc *ptrNextSubFragLoc;
     SubFragLoc();
	 SubFragLoc(unsigned _int32 intArray_Size);
     ~SubFragLoc();
};

//Bond ID Bit Field
class BondID {
public:
     unsigned int Bond_Type:4;  // 1~single,2~double,3~triple,4~aromatic,15~wild
     unsigned int Not_Bond:1;
     BondID();
     ~BondID();
};

//Class representing the bonding info and identity of the other atom in the bond.
class AtomBond {
public:
     BondID Bond_ID;
     unsigned _int32 intAttachedAtom;
     AtomBond *ptrNextBond;
     void AddConnection(unsigned _int8 intBondValue,unsigned _int32 intPreviousAtom,unsigned _int8 intNotBond);
     AtomBond();
     ~AtomBond();
};

//Atom ID Bit Field
class AtomID {
public:
     unsigned int Atom_Type:7;
     unsigned int Num_NH_Cnt:3;
     unsigned int Atom_Charge:2;
     unsigned int Bond_Types:4;         //  | 3~aryl | 2~triple | 1~double | 0~single |
     unsigned int Atom_Cnt:5;           //  |4~ misc. | 3~halogens | 2~O,S,Se | 1~N,P,B | 0~C,Si |
     AtomID();
     ~AtomID();
};

//Search ID Bit Field
class SearchID {
public:
     //Ring Search Codes
	 unsigned int Ring_Size:4;
     unsigned int Less_Than:1;
     unsigned int Greater_Than:1;
     unsigned int Not_Ring:1;
     unsigned int In_Ring:1;
     unsigned int Ring_Type:2;          // | 1~aryl | 0~alicyclic |
     //Atom Search Codes
     unsigned int Unspec_Bond:1;
     unsigned int Unspec_Atom:1;
     unsigned int Used_Atom:1;
     unsigned int Dechar_Atom:1;
     unsigned int Not_Atom:1;
     unsigned int Unspec_Neighbor:1;

	 void Reset_ID();
     SearchID();
     ~SearchID();
};

//Class containing the processing flags for ChemSeqID.
class ProcessFlags {
public:
	unsigned int MF_Fill:1;		//Denotes that the molecular feature bitfield has been processed.
	unsigned int RSC_Fill:1;	//Denotes that the ring search code bitfield in the Search_ID has been processed.
	unsigned int SSSR_Fill:1;	//Denotes that ring perception has been performed on the molecule.
	unsigned int QK_Fill:1;		//Denotes that the quantity key has been processed.

	ProcessFlags();
	~ProcessFlags();
};

//Class representing the atom (node) in molecular structure.  Atom (node) 
//information is specified in the characteristic bitfield.
class Atom {
public:
     AtomID Atom_ID;
     SearchID Search_ID;
     void AddConnection(unsigned _int8 intBondValue,unsigned _int32 intPreviousAtom,unsigned _int8 intNotBond);
     AtomBond *ptrNextBond;
     Atom();
     ~Atom();
};

//Molecular Feature Bit Field
class MolecularFeature {
public:
     //Ring Features
	 unsigned int Ring_Feature:1;       //Denotes MF was a ring feature.
	 unsigned int Ring_Size:4;			//up to 15 atoms
     unsigned int R_Less_Than:1;
     unsigned int R_Greater_Than:1;
     unsigned int Not_Ring:1;
     unsigned int Ring_Type:3;          //  | 1~aryl | 0~alicyclic | 
     
	 //Bond Features
     unsigned int Bond_Feature:1;		//Denotes MF was a bond feature.
	 unsigned int B_Less_Than:1;
     unsigned int B_Greater_Than:1;
     unsigned int Not_Bond:1;
     unsigned int Bond_Type:3;			//Used as integer 7~all,6-5~blank,4~aryl,3~triple,2~double,1~single,0~ionic
	 unsigned int Bond_Constraint:3;	//Used as bitfield /2~blank/1~in ring/0~not in ring/
     
	 MolecularFeature();
     ~MolecularFeature();
};

//Bitfield used to store number of atoms of specified type for group contribution type screening
//(using a 96 bit bitfield for screening ID)
class QK_BF {
public:
     unsigned int Num_C:8;		//up to 255 aliphatic carbon atoms
     unsigned int Num_O:6;		//up to 63 aliphatic oxygen atoms
     unsigned int Num_N:5;		//up to 31 aliphatic nitrogen atoms
     unsigned int Num_S:5;		//up to 31 aliphatic sulfur atoms
     unsigned int Num_P:3;		//up to 7 phosphorous atoms
     unsigned int Num_Br:4;		//up to 15 bromine atoms
     unsigned int Num_F:4;		//up to 15 fluorine atoms
     unsigned int Num_Cl:4;		//up to 15 chlorine atoms
     unsigned int Num_I:4;		//up to 15 iodine atoms
     unsigned int Num_Etc:4;	//up to 15 miscellaneous (i.e., Si,Se,B,etc.) atoms
     unsigned int Num_DB:7;		//up to 127 double bonds
     unsigned int Num_TB:4;		//up to 15 triple bonds
     unsigned int Num_CDBO:5;	//up to 31 carbon double bonded to oxygen atoms
	 unsigned int Num_CDBC:7;	//up to 127 carbon double bonded to carbon atoms
	 unsigned int Num_CTBN:3;	//up to 7 carbon triple bonded to nitrogen atoms
     unsigned int Num_AC:7;		//up to 127 aromatic carbon atoms
     unsigned int Num_AO:3;		//up to 7 aromatic oxygen atoms
     unsigned int Num_AS:3;		//up to 7 aromatic sulfur atoms
     unsigned int Num_AN:3;		//up to 7 aromatic nitrogen atoms
     unsigned int Num_Rings:6;	//up to 63 rings in structure
     
	 //Screen ID functions
	 void Copy_QK(QK_BF &Tmp_QK);
	 void Truncate_QK(SubFragLoc *ptrAlpha_Loc,Atom *ptrAlpha_Molecule,Atom *ptrBeta_Molecule,unsigned _int32 intNum_Beta_Atoms);
	 void Reset_QK();
	 
	 QK_BF();
	 ~QK_BF();
};

//Class representing the molecular graph (points to first atom)
class Molecule {
public:
	Atom *ptrAtom;

	Molecule();
	~Molecule();
};

//Class storing the detected "smallest set of smallest rings" for each query structure.
class SSSR {
public:
	unsigned int Ring_Type:4;		// | 1~aromatic | 0~alicyclic |
	unsigned int Num_Members:10;	//Number of ring members.
	unsigned _int32 *ptrRing_Mem;	//Array of member locations in query structure.
	SSSR *ptrNextSSSR;

	SSSR(unsigned _int32 intType,unsigned _int32 intNum_Atoms,unsigned _int32 *ptrRing_Set,unsigned _int32 intArray_Length);
	SSSR();
	~SSSR();
};

//Class representing a chemical identification node for sequential searching
class ChemSeqID {
private:
	 //GRAPH MEMBER FUNCTIONS
	 //Figueras_SSSR() internal "node trimming" routine.
	 inline unsigned _int8 Intersect_Path(unsigned _int32 **ptrAtom_Path,unsigned _int32 intRoot_Node,unsigned _int32 intTop_Node,unsigned _int32 intCurrent_Node,unsigned _int32 intArray_Length);
	 inline void Append_Paths(unsigned _int32 *ptrTarget_Path,unsigned _int32 *ptrAdd_Path1,unsigned _int32 *ptrAdd_Path2,unsigned _int32 intArray_Length);
	 inline void Return_SSSR(unsigned _int32 *ptrRing_Set,unsigned _int32 *ptrAtom_Path1,unsigned _int32 *ptrAtom_Path2,unsigned _int32 intArray_Length,unsigned _int32 &intNum_Ring_Atoms);
	 inline unsigned _int8 Check_Path(unsigned _int32 **ptrAtom_Path,unsigned _int32 intAtom_Loc,unsigned _int32 intArray_Length);
	 inline unsigned _int8 ChemSeqID::Compare_Set(unsigned _int32 *ptrSet1,unsigned _int32 *ptrSet2,unsigned _int32 intArray_Length);
	 inline unsigned _int8 Check_Duplicate(unsigned _int32 *ptrRing_Set,unsigned _int32 intArray_Length);
	 inline void Un_Trim_SSSR(unsigned _int32 intTrim_Loc,unsigned _int8 *ptrDegree,unsigned _int32 *ptrTrim_Set);
	 inline void Trim_SSSR(unsigned _int32 intTrim_Loc,unsigned _int8 *ptrDegree,unsigned _int32 *ptrTrim_Set);
	 inline unsigned _int8 Compare_Set(unsigned _int32 *ptrFull_Set,unsigned _int32 *ptrTrim_Set);
	 inline void Calc_Ring_Type(SSSR *ptrTemp_SSSR);
	 void Check_Nodes(unsigned _int32 intRoot_Node,unsigned _int32 intRing_Size,unsigned _int32 *ptrRing_Set,unsigned _int32 *ptrTrim_Set,unsigned _int8 *ptrDegree,unsigned _int32 intArray_Length);
	 unsigned _int8 Get_Ring(unsigned _int32 intRoot_Node,unsigned _int32 *ptrRing_Set,unsigned _int8 *ptrDegree,unsigned _int32 &intRing_Size,unsigned _int32 intArray_Length);

public:
 	 //GRAPH MEMBER FUNCTIONS
	 void Set_SOF_BF(unsigned _int8 &intOF_Cntr,unsigned _int8 &intOF_BF,unsigned _int8 &intGF_BF,unsigned _int8 *intOF_Que,Atom &ptrAtom,unsigned _int32 intCharCount,SearchID Temp_SID);
	 void AttachNewAtom(Atom *ptrStructure,unsigned _int32 intNew_Loc,AtomBond **ptrBranchAttach,_int32 *intBranchLoc,unsigned _int32 intAttachCount,unsigned _int8 &intBondValue,unsigned _int8 intNotBond);
	 void Fill_AID_SID();
	 void Calc_RSC();	//Fill ring search codes in Search_ID using graph algorithms.
	 unsigned _int16 Retrieve_Bond_MF(ChemSeqID *ptrBetaIDNode);
	 unsigned _int16 Retrieve_Ring_MF(ChemSeqID *ptrBetaIDNode);
	 unsigned _int8 Figueras_SSSR();

	 //SMILES MEMBER FUNCTIONS
     unsigned _int16 Get_SMILES_RN(unsigned _int32 &intCharCount);	      
	 void SMILES_QK_Fill();
	 void Parse_SMILES();

	 //SCREEN & SEARCH SYNTAX ROUTINES
	 void QK_to_OK();	
	 void Fill_MF();

	 //DATA MEMBERS
	 char *chrChemSyntaxString;			//String representing chemical structure.
     ChemSeqID *ptrNextSubFrag;		
      _int32 intChemEntryID;			//Integer ID used to denote chemical structure.
     unsigned _int16 intOccupancyKey;	//Key denoting that specific atomic keys are present in structure (file searching).
     unsigned _int32 intNumberAtoms;	//Number of non hydrogen atoms in chemical structure.
     QK_BF intQuantKey;					//Key denoting quantity of specific atomic keys present. Used in pre-screening.
     MolecularFeature MF_ID;			//Molecular feature bitfield used int MF searching.
	 ProcessFlags Flags;				//Processing flags bitfield.
	 Molecule *ptrMolecule;				//Pointer to molecular structure (graph).
	 SSSR *ptrSSSR;						//Stores smallest set of smallest rings detection.
	 
	 ChemSeqID();
     ~ChemSeqID();
};

//Class storing the detected beta locations in the alpha graph in bit field integer arrays
class ComboLoc {
public:
     ChemSeqID *ptrChemID;
     unsigned _int32 *ptrQuery_Loc;
     ComboLoc *ptrNextComboLoc;
     ComboLoc *ptrNextDisjoint;
	 ComboLoc();
     ~ComboLoc();
};

//Class representing a chemical identification node for sequential searching
class ChemComboID:public ChemSeqID {
private:
	//Sorting routine to order the subfragment detections in ascending order within a cover.
 	void Bounded_QuickSort(ComboLoc **Key_Vector,unsigned _int32 *Slave_Vector,ComboLoc *ptrMax_Bound,unsigned _int32 intLower_Bounds,unsigned _int32 intLow,unsigned _int32 intUpper_Bounds,unsigned _int32 intHigh);
public:
	unsigned _int32 *ptrEC_Check;
	unsigned _int32 intEC_Array_Length;
	ComboLoc *ptrFirstComboLoc;
     
	ChemComboID();
    ~ChemComboID();
	
	//Routine to determine whether a given cover is a degenerate occurrence of previously detected covers.
	unsigned _int8 Check_Degeneracy(ComboLoc **ptrFragID,unsigned _int32 *ptrBitCntr,unsigned _int32 &intDetectionCntr,unsigned _int32 &intOldDemarcation,unsigned _int32 intNum_Groups[30],unsigned _int32 &intGroup_Cntr);
	//EC_Check array routines.
	void Initialize_EC_Check();
	void Fill_EC_Check(ComboLoc *ptrTempComboLoc);
	unsigned _int8 Compare_EC_Check();
};

//Class representing a list of molecule strings for a given input file
class SubFragList {
public:
     ChemSeqID *ptrFirstIDNode;				//Pointer to first subfragment in subfragment list
     ChemSeqID *(*ptrBinSortLoc);
	 unsigned _int32 intAvg_Num_Atoms;		//Average number of atoms
	 unsigned _int16 intNumber_SF;			//Number of subfragments
     
	 void ListBinSort(unsigned _int8 intNumPockets);
     unsigned _int8 Read_SF_List(char *strSubFragFile,_int8 intQueryType);
	 SubFragList();
     ~SubFragList();
};

#endif    //H_SDA_1