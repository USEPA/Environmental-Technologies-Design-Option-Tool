///////////////////////////////////////////////////////////////
//	This is the constructor/destructor code accompanying the 
//	header file SDA_1.h for MOSDAP(c) v2.0
//
//	Copyright (c) 1998. John W. Raymond.  All rights reserved.
//
///////////////////////////////////////////////////////////////

#include "SDA_1.H"

SubFragLoc::SubFragLoc() {
     ptrQuery_Loc=0;
	 intArray_Length=0;
     ptrNextSubFragLoc=0;
}

SubFragLoc::SubFragLoc(unsigned _int32 intArray_Size) {
     unsigned _int32 i;
	 intArray_Length=intArray_Size;

     //Create an array of length intArray_Length and initialize to zero.
     ptrQuery_Loc=new unsigned _int32[intArray_Length];
     for(i=0;i<intArray_Length;i++) {
          ptrQuery_Loc[i]=0;
     }
     ptrNextSubFragLoc=0;
}

SubFragLoc::~SubFragLoc() {
     if(ptrQuery_Loc) delete[] ptrQuery_Loc;
     if(ptrNextSubFragLoc) delete ptrNextSubFragLoc;
}

BondID::BondID() {
     Bond_Type=0;
     Not_Bond=0;
}

BondID::~BondID() {

}

AtomBond::AtomBond() {
     intAttachedAtom=0;
     ptrNextBond=0;
}

AtomBond::~AtomBond() {
     if(ptrNextBond) delete ptrNextBond;
}

AtomID::AtomID() {
     Atom_Type=0;
     Num_NH_Cnt=0;
     Atom_Charge=0;
     Bond_Types=0;
     Atom_Cnt=0;
}

AtomID::~AtomID() {

}

SearchID::SearchID() {
	 Ring_Size=0;
     Less_Than=0;
     Greater_Than=0;
     Not_Ring=0;
     In_Ring=0;
     Ring_Type=0;
     
	 Unspec_Bond=0;
     Unspec_Atom=0;
     Used_Atom=0;
     Dechar_Atom=0;
     Not_Atom=0;
     Unspec_Neighbor=0;
}

SearchID::~SearchID() {

}

ProcessFlags::ProcessFlags() {
	MF_Fill=0;
	RSC_Fill=0;
	SSSR_Fill=0;
	QK_Fill=0;
}

ProcessFlags::~ProcessFlags() {

}

Atom::Atom() {
     ptrNextBond=0;
}

Atom::~Atom() {
     if(ptrNextBond) delete ptrNextBond;
}

MolecularFeature::MolecularFeature() {
     Ring_Feature=0;
	 Ring_Size=0;
     R_Less_Than=0;
     R_Greater_Than=0;
     Not_Ring=0;
     Ring_Type=0;

	 Bond_Feature=0;
     B_Less_Than=0;
     B_Greater_Than=0;
     Not_Bond=0;
	 Bond_Type=0;
	 Bond_Constraint=0;
}

MolecularFeature::~MolecularFeature() {

}

QK_BF::QK_BF () {

     Num_C=0;
     Num_O=0;
     Num_N=0;
     Num_S=0;
     Num_P=0;
     Num_Br=0;
     Num_F=0;
     Num_Cl=0;
     Num_I=0;
     Num_Etc=0;
     Num_DB=0;
     Num_TB=0;
     Num_CDBO=0;
	 Num_CDBC=0;
	 Num_CTBN=0;
     Num_AC=0;
     Num_AO=0;
     Num_AS=0;
     Num_AN=0;
     Num_Rings=0;
}

QK_BF::~QK_BF() {

}

Molecule::Molecule() {
	ptrAtom=0;
}

Molecule::~Molecule() {
	 if(ptrAtom) delete[] ptrAtom;
}

SSSR::SSSR(unsigned _int32 intType,unsigned _int32 intNum_Atoms,unsigned _int32 *ptrRing_Set,unsigned _int32 intArray_Length){

	unsigned _int32 i;

	Ring_Type=intType;
	Num_Members=intNum_Atoms;
	ptrRing_Mem=new unsigned _int32[intArray_Length];
	
	for(i=0;i<intArray_Length;i++) {
		ptrRing_Mem[i]=ptrRing_Set[i];
	}
	ptrNextSSSR=0;

}

SSSR::SSSR() {
	Ring_Type=0;
	Num_Members=0;
	ptrRing_Mem=0;
	ptrNextSSSR=0;
}

SSSR::~SSSR() {
	if(ptrRing_Mem) delete[] ptrRing_Mem;
	if(ptrNextSSSR) delete ptrNextSSSR;
}

ChemSeqID::ChemSeqID() {
     chrChemSyntaxString=0;
     ptrNextSubFrag=0;
     intChemEntryID=0;
     intOccupancyKey=0;
     intNumberAtoms=0;
     ptrMolecule=0;
	 ptrSSSR=0;
}

ChemSeqID::~ChemSeqID() {
     if(chrChemSyntaxString) delete[] chrChemSyntaxString;
     if(ptrMolecule) delete ptrMolecule;
	 if(ptrNextSubFrag) delete ptrNextSubFrag;
	 if(ptrSSSR) delete ptrSSSR;
}

ComboLoc::ComboLoc() {
     ptrChemID=0;
     ptrQuery_Loc=0;
     ptrNextComboLoc=0;
     ptrNextDisjoint=0;
}

ComboLoc::~ComboLoc() {
     if(ptrChemID) ptrChemID=0;
     if(ptrQuery_Loc) delete[] ptrQuery_Loc;
     if(ptrNextDisjoint) ptrNextDisjoint=0;
	 if(ptrNextComboLoc) delete ptrNextComboLoc;
}

ChemComboID::ChemComboID() {
    ptrEC_Check=0; 
	ptrFirstComboLoc=0;
	intEC_Array_Length=0;
}

ChemComboID::~ChemComboID(){
     if(ptrFirstComboLoc) delete ptrFirstComboLoc;
	 if(ptrEC_Check) delete ptrEC_Check;
}

SubFragList::SubFragList() {
     ptrFirstIDNode=0;
     ptrBinSortLoc=0;
	 intAvg_Num_Atoms=0;
	 intNumber_SF=0;
}

SubFragList::~SubFragList() {
     if(ptrFirstIDNode) delete ptrFirstIDNode;
     if(ptrBinSortLoc) delete[] ptrBinSortLoc;
}


