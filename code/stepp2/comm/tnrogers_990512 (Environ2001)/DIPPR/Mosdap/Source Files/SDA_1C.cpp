///////////////////////////////////////////////////////////////
//	This is the function code accompanying the 
//	header file SDA_1.h for MOSDAP(c) v2.0
//
//	Copyright (c) 1998. John W. Raymond.  All rights reserved.
//
///////////////////////////////////////////////////////////////

#include "SDA_1.H"
#include "SDA_2.H"

void
ChemSeqID::Set_SOF_BF(unsigned _int8 &intOF_Cntr,unsigned _int8 &intOF_BF,unsigned _int8 &intGF_BF,unsigned _int8 *intOF_Que,Atom &ptrAtom,unsigned _int32 intCharCount,SearchID Temp_SID) {
//intOF_BF => Flag bit field for molecular operators  ...| 3~MF operator | 2~Not atom | 1~Dechar atom |0~blank | 

//If the decharacterization flag is set then set the decharacterization bit.
if(intOF_BF & (1<<1)) ptrAtom.Search_ID.Dechar_Atom=1;
//If the not flag is set then set the not atom flag.
if(intOF_BF & (1<<2)) ptrAtom.Search_ID.Not_Atom=1;
//If the MF operator flag is set then set the appropriate MF flags in the search ID
if(intOF_BF & (1<<3)) {
     ptrAtom.Search_ID.Ring_Size=Temp_SID.Ring_Size;
     ptrAtom.Search_ID.Less_Than=Temp_SID.Less_Than;
     ptrAtom.Search_ID.Greater_Than=Temp_SID.Greater_Than;
     ptrAtom.Search_ID.Not_Ring=Temp_SID.Not_Ring;
     ptrAtom.Search_ID.In_Ring=Temp_SID.In_Ring;
     ptrAtom.Search_ID.Ring_Type=Temp_SID.Ring_Type;
}

//Reset Operator Flags for single atom occurrences (no group flag set).
//Grouping braces {,} are associated w/ at least one of the current operators.
if(intGF_BF) {
	while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))) {
		//If despecification operator is associated with this combo bracket, then set back operator flags.
        if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
			intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
            intOF_Que[intOF_Cntr]=0;
            if(intOF_Cntr)intOF_Cntr--;
        }
	}
}
//No grouping braces {,} are associated w/ current operators.
else {
	intOF_BF=0;
	intOF_Cntr=0;
	intOF_Que[intOF_Cntr]=0;
}

}

//Routine used in Parse_SMILES to add a new atom and its associated connections to a molecular graph.
void 
ChemSeqID::AttachNewAtom(Atom *ptrStructure,unsigned _int32 intNew_Loc,AtomBond **ptrBranchAttach,_int32 *intBranchLoc,unsigned _int32 intAttachCount,unsigned _int8 &intBondValue,unsigned _int8 intNotBond) {

AtomBond *ptrTempAtom;

//If not first atom in string, then add current connection to previous attached connection node at connection anchor.
if(ptrBranchAttach[intAttachCount]) {
	while(ptrBranchAttach[intAttachCount]->ptrNextBond) {
		ptrBranchAttach[intAttachCount]=ptrBranchAttach[intAttachCount]->ptrNextBond;
	}
	ptrBranchAttach[intAttachCount]->AddConnection(intBondValue,intNew_Loc,intNotBond);
}
//Else add node to adjacency listing spine.
else {
	ptrStructure[0].AddConnection(intBondValue,intNew_Loc,intNotBond);
	ptrBranchAttach[0]=ptrStructure[0].ptrNextBond;
}

//Add current connection to current atom.
ptrTempAtom=ptrStructure[intNew_Loc].ptrNextBond;
//New node has previous branches (i.e., in a ring)
if(ptrTempAtom) {
     while(ptrTempAtom->ptrNextBond) {
          ptrTempAtom=ptrTempAtom->ptrNextBond;
     }
     ptrTempAtom->AddConnection(intBondValue,intBranchLoc[intAttachCount],intNotBond);
     ptrBranchAttach[intAttachCount]=ptrTempAtom->ptrNextBond;
}
//New node has no previous branches.
else {
     ptrStructure[intNew_Loc].AddConnection(intBondValue,intBranchLoc[intAttachCount],intNotBond);
     ptrBranchAttach[intAttachCount]=ptrStructure[intNew_Loc].ptrNextBond;
}

intBranchLoc[intAttachCount]=intNew_Loc;
intBondValue=1;
intNotBond=0;
}

//Routine to fill ring search codes in search ID for each atom in structure.
void
ChemSeqID::Calc_RSC(){

	unsigned _int32 i;						//Temporary counter
	const unsigned _int8 intBit_Size=32;	//Size of integer used in storage arrays
	SSSR *ptrTemp_SSSR;						//Temporary SSSR pointer

	//If a ring was detected in the structure.
	if(Figueras_SSSR()) {

		ptrTemp_SSSR=ptrSSSR;

		while(ptrTemp_SSSR) {

			for(i=0;i<intNumberAtoms;i++) {

				//Flag atom as processed for ring analysis.
				Flags.RSC_Fill=1;

				if(ptrTemp_SSSR->ptrRing_Mem[i/intBit_Size] & (1<<(1%intBit_Size))) {

					ptrMolecule->ptrAtom[i].Search_ID.Ring_Type=ptrTemp_SSSR->Ring_Type;
					ptrMolecule->ptrAtom[i].Search_ID.In_Ring=1;
					ptrMolecule->ptrAtom[i].Search_ID.Ring_Size=ptrTemp_SSSR->Num_Members;
				}
				else ptrMolecule->ptrAtom[i].Search_ID.Not_Ring=1;
			}
			ptrTemp_SSSR=ptrTemp_SSSR->ptrNextSSSR;
		}
	}
}

//Routine to determine the ring type for a detected ring.  This is where the aromaticity detection algorithm for a detected ring (Kekule structure) would
//be placed. Currently, the routine only supports syntax declared aromatic and alicyclic properties.  Future versions can incorporated declarations of fused
//rings, hetero-atom rings, etc.
inline void
ChemSeqID::Calc_Ring_Type(SSSR *ptrTemp_SSSR) {

	const unsigned _int8 intBit_Size=32;	//Size of integers.
	unsigned _int32 i;

	for(i=0;i<intNumberAtoms;i++) {

		//Find an atom in the detected ring.
		if(ptrTemp_SSSR->ptrRing_Mem[i/intBit_Size]& (1<<(i%intBit_Size))) {

			//If atom has been syntax declared previously as being in an aromatic ring.
			if(ptrMolecule->ptrAtom[i].Search_ID.In_Ring && (ptrMolecule->ptrAtom[i].Search_ID.Ring_Type & (1<<1))){
				//Set as aryl.
				ptrTemp_SSSR->Ring_Type |=(1<<1);
				break;
			}
		}
	}
	//Else if ring has not been determined to be aromatic, then flag as alicyclic.
	if(!(ptrTemp_SSSR->Ring_Type & (1<<1))) ptrTemp_SSSR->Ring_Type |= (1<<0);

	return;
}

//Find intersection of ring paths. Returns 1 if intersection is a singleton or 0 if intersection is not a singleton.
inline unsigned _int8
ChemSeqID::Intersect_Path(unsigned _int32 **ptrAtom_Path,unsigned _int32 intRoot_Node,unsigned _int32 intTop_Node,unsigned _int32 intCurrent_Node,unsigned _int32 intArray_Length){

const unsigned _int8 intBit_Size=32;
unsigned _int32 intIntersect_Place;
unsigned _int32 i;

intIntersect_Place=intRoot_Node/intBit_Size;

for(i=0;i<intArray_Length;i++) {

	if(i==intIntersect_Place) {

		if((ptrAtom_Path[intTop_Node][i] & ptrAtom_Path[intCurrent_Node][i]) ^ (1<<(intRoot_Node%intBit_Size))) return 0;
	}
	else {
		if(ptrAtom_Path[intTop_Node][i] & ptrAtom_Path[intCurrent_Node][i]) return 0;
	}
}	
return 1;

}

//Appends two paths to a target path set for Figueras algorithm.
inline void
ChemSeqID::Append_Paths(unsigned _int32 *ptrTarget_Path,unsigned _int32 *ptrAdd_Path1,unsigned _int32 *ptrAdd_Path2,unsigned _int32 intArray_Length) {

	unsigned _int32 i;

	for(i=0;i<intArray_Length;i++) {
		ptrTarget_Path[i] |= (ptrAdd_Path1[i] | ptrAdd_Path2[i]);
	}
}

//Returns results of detected ring structure (path and number member elements).
inline void
ChemSeqID::Return_SSSR(unsigned _int32 *ptrRing_Set,unsigned _int32 *ptrAtom_Path1,unsigned _int32 *ptrAtom_Path2,unsigned _int32 intArray_Length,unsigned _int32 &intNum_Ring_Atoms) {

	unsigned _int32 i;
	const unsigned _int8 intBit_Size=32;	//Size of array integers in bits.

	intNum_Ring_Atoms=0;

	Append_Paths(ptrRing_Set,ptrAtom_Path1,ptrAtom_Path2,intArray_Length);

	for(i=0;i<intNumberAtoms;i++) {
		if(ptrRing_Set[i/intBit_Size] & (1<<(i%intBit_Size))) intNum_Ring_Atoms++;
	}

}

//Checks if ring path for Figueras (1996) search is null. Returns 1 if set is non-null and zero if set is null.
inline unsigned _int8
ChemSeqID::Check_Path(unsigned _int32 **ptrAtom_Path,unsigned _int32 intAtom_Loc,unsigned _int32 intArray_Length) {

	unsigned _int32 i;

	for(i=0;i<intArray_Length;i++) {
		if(ptrAtom_Path[intAtom_Loc][i]) return 1;
	}
	return 0;

}

//Routine used int Figueras_SSSR() to compare the an atom set with another set to determine if the sets are equal.
//Return 1 if they are equal and 0 if they are not equal.
inline unsigned _int8
ChemSeqID::Compare_Set(unsigned _int32 *ptrSet1,unsigned _int32 *ptrSet2,unsigned _int32 intArray_Length) {

	unsigned _int32 i;

	for(i=0;i<intArray_Length;i++) {
		if(ptrSet1[i] != ptrSet2[i]) return 0;
	}
	return 1;

}

//Checks if detected ring is a duplicate detection for Figueras Ring detection algorithm (1996).
//Returns 1 if it is not a duplicate detection, else it returns 0.
inline unsigned _int8
ChemSeqID::Check_Duplicate(unsigned _int32 *ptrRing_Set,unsigned _int32 intArray_Length) {

	SSSR *ptrTemp1_SSSR;

	ptrTemp1_SSSR=ptrSSSR;

	while(ptrTemp1_SSSR) {

		if(Compare_Set(ptrTemp1_SSSR->ptrRing_Mem,ptrRing_Set,intArray_Length)) return 0;
		
		ptrTemp1_SSSR=ptrTemp1_SSSR->ptrNextSSSR;
	}
	return 1;
}

//Routine used in Figueras_SSSR() to un-trim a node from the query structure for further processing.
inline void
ChemSeqID::Un_Trim_SSSR(unsigned _int32 intTrim_Loc,unsigned _int8 *ptrDegree,unsigned _int32 *ptrTrim_Set) {

	const unsigned _int8 intBit_Size=32;	//Bit size of integer used in Trim_Set[] array.
	unsigned _int32 intBit_Loc;				//Bit location in array integer of current node.
	unsigned _int32 intArray_Loc;			//Array location in array of current node.
	AtomBond *ptrTemp_Bond;					//Bond pointer used to loop through neighbors of current node.

	//Reset "trimmed node" bit in trimSet.
	intArray_Loc=intTrim_Loc/intBit_Size;
	intBit_Loc=intTrim_Loc%intBit_Size;
	ptrTrim_Set[intArray_Loc] ^= (1<<intBit_Loc);

	//Decrement the degree of "trimmed node's" neighbors.
	ptrTemp_Bond=ptrMolecule->ptrAtom[intTrim_Loc].ptrNextBond;
	while(ptrTemp_Bond) {
		if(ptrDegree[ptrTemp_Bond->intAttachedAtom])ptrDegree[ptrTemp_Bond->intAttachedAtom]++;
		
		//Increment temporarily "trimmed node" degree.
		ptrDegree[intTrim_Loc]++;

		ptrTemp_Bond=ptrTemp_Bond->ptrNextBond;
	}
}

//Routine to find all rings which an N3 node elimination candidate is a member and then eliminate the node,
//thereby, creating new N2 nodes for subsequent ring search.
void
ChemSeqID::Check_Nodes(unsigned _int32 intRoot_Node,unsigned _int32 intRing_Size,unsigned _int32 *ptrRing_Set,unsigned _int32 *ptrTrim_Set,unsigned _int8 *ptrDegree,unsigned _int32 intArray_Length) {

	unsigned _int32 intN2_Cntr;
	unsigned _int32 intMin_Size=0;		//Minimum ring size (number of members)
	unsigned _int32 intMin_Place=0;		//Location 
	unsigned _int32 intRing_Type=0;		//Code determining the chemical nature of the detected ring.
	unsigned _int32 *ptrNodes_N2;		//Array storing temporarily created N2 nodes.
	unsigned _int32 i;
	unsigned _int32 j;
	const unsigned _int8 intBit_Size=32;
	SSSR *ptrTemp_SSSR;
	if(ptrSSSR) ptrTemp_SSSR=ptrSSSR;

	ptrNodes_N2=new unsigned _int32[intNumberAtoms];

	for(i=0;i<intNumberAtoms;i++) {

		//Find a node member of the current detected ring.
		if(ptrRing_Set[i/intBit_Size] & (1<<i%intBit_Size)) {

			//Temporarily trim the current ring node from the ring.
			Trim_SSSR(i,ptrDegree,ptrTrim_Set);

			intN2_Cntr=0;
			//Fill nodes N2 array.
			for(j=0;j<intNumberAtoms;j++) {
				if(ptrDegree[j]==2) {
					ptrNodes_N2[intN2_Cntr]=j;
					intN2_Cntr++;
				}
			}

			//Call get_ring() routine for each temporary N2 node.
			for(j=0;j<intN2_Cntr;j++) {

				if(Get_Ring(ptrNodes_N2[j],ptrRing_Set,ptrDegree,intRing_Size,intArray_Length)) {

					//Check for duplicate rings, if not duplicate then add to SSSR.
					if(!ptrSSSR) {
						ptrSSSR=new SSSR(intRing_Type,intRing_Size,ptrRing_Set,intArray_Length);
						ptrTemp_SSSR=ptrSSSR;
					}
					else if(Check_Duplicate(ptrRing_Set,intArray_Length)) {
						//Go to end of SSSR list.
						while(ptrTemp_SSSR->ptrNextSSSR) {
							ptrTemp_SSSR=ptrTemp_SSSR->ptrNextSSSR;
						}
						ptrTemp_SSSR->ptrNextSSSR=new SSSR(intRing_Type,intRing_Size,ptrRing_Set,intArray_Length);
						ptrTemp_SSSR=ptrTemp_SSSR->ptrNextSSSR;
					}

					//Determine the ring type of detected ring.
					Calc_Ring_Type(ptrTemp_SSSR);
					
					//Determine which node in orignial ring detection is a member of the smallest ring.
					if(!intMin_Size || intRing_Size<intMin_Size) {
						intMin_Size=intRing_Size;
						intMin_Place=i;
					}

				}
			}

			//Reset the temporarily trimmed N3 node.
			Un_Trim_SSSR(i,ptrDegree,ptrTrim_Set);
		}
	}
	//Permanently trim N3 node resulting in smallest ring detection.
	Trim_SSSR(intMin_Place,ptrDegree,ptrTrim_Set);

	delete ptrNodes_N2;
}

//Routine to perform BFS on array of nodes of degree two to detect a potential ring.
unsigned _int8
ChemSeqID::Get_Ring(unsigned _int32 intRoot_Node,unsigned _int32 *ptrRing_Set,unsigned _int8 *ptrDegree,unsigned _int32 &intRing_Size,unsigned _int32 intArray_Length) {

unsigned _int32 i,j;				//Junk variables
unsigned _int32 intTop_Node;		//Location of top node in que
unsigned _int32 **ptrAtom_Path;		//Array for each node representing its cumulative path
const unsigned _int8 intBit_Size=32;//Size of array integers

QueNode	*ptrTemp1_Que,*ptrTemp2_Que,*ptrTemp3_Que;	//Que pointer
AtomBond *ptrTemp_Bond;								//Bond pointer

//Initialize arrays.
ptrAtom_Path=new unsigned _int32*[intNumberAtoms];

for(i=0;i<intNumberAtoms;i++) {
	ptrAtom_Path[i]=new unsigned _int32[intArray_Length];
	for(j=0;j<intArray_Length;j++) {
		ptrAtom_Path[i][j]=0;
		ptrRing_Set[j]=0;
	}
}

//Initialize the queue with nodes attached to the root node.
ptrTemp_Bond=ptrMolecule->ptrAtom[intRoot_Node].ptrNextBond;
ptrTemp2_Que=new QueNode;
ptrTemp1_Que=ptrTemp2_Que;

while(ptrTemp_Bond) {
	if(ptrDegree[ptrTemp_Bond->intAttachedAtom]>0) {

		if(ptrTemp_Bond != ptrMolecule->ptrAtom[intRoot_Node].ptrNextBond) {
			ptrTemp2_Que->Add_Que(intRoot_Node,ptrTemp_Bond->intAttachedAtom);
			ptrTemp2_Que=ptrTemp2_Que->ptrNextQue;
		}
		else {
			ptrTemp2_Que->intSource=intRoot_Node;
			ptrTemp2_Que->intAtom_Index=ptrTemp_Bond->intAttachedAtom;
		}
		
		ptrAtom_Path[ptrTemp_Bond->intAttachedAtom][intRoot_Node/intBit_Size] |= (1<<(intRoot_Node%intBit_Size));
		ptrAtom_Path[ptrTemp_Bond->intAttachedAtom][ptrTemp_Bond->intAttachedAtom/intBit_Size] |= (1<<(ptrTemp_Bond->intAttachedAtom%intBit_Size));

	}
	ptrTemp_Bond=ptrTemp_Bond->ptrNextBond;
}

//Examine Que
while(ptrTemp1_Que) {

	//Get data from the first que node.
	intTop_Node=ptrTemp1_Que->intAtom_Index;

	ptrTemp_Bond=ptrMolecule->ptrAtom[intTop_Node].ptrNextBond;
	while(ptrTemp_Bond) {

		//Examine nodes attached to the top node. Avoid the top node's source.
		if(ptrDegree[ptrTemp_Bond->intAttachedAtom]>0 && ptrTemp_Bond->intAttachedAtom != ptrTemp1_Que->intSource) {

			//Collision occurs when attached node's path is no longer empty.
			if(Check_Path(ptrAtom_Path,ptrTemp_Bond->intAttachedAtom,intArray_Length)) {

				//If intersection is a singleton.
				if(Intersect_Path(ptrAtom_Path,intRoot_Node,intTop_Node,ptrTemp_Bond->intAttachedAtom,intArray_Length)) {

					//Return results (ring path and number atoms in ring) to the SSSR routine.
					Return_SSSR(ptrRing_Set,ptrAtom_Path[intTop_Node],ptrAtom_Path[ptrTemp_Bond->intAttachedAtom],intArray_Length,intRing_Size);

					//Clean up dynamic memory allocations.
					for(j=0;j<intNumberAtoms;j++) {
						delete[] ptrAtom_Path[j];
					}
					delete[] ptrAtom_Path;
					if(ptrTemp1_Que) delete ptrTemp1_Que;
					return 1;

				}
			}
			else {
				//Update the path ptrTemp_Bond->intAttachedAtom
				ptrAtom_Path[ptrTemp_Bond->intAttachedAtom][ptrTemp_Bond->intAttachedAtom/intBit_Size] |= (1<<(ptrTemp_Bond->intAttachedAtom%intBit_Size));
				Append_Paths(ptrAtom_Path[ptrTemp_Bond->intAttachedAtom],ptrAtom_Path[intTop_Node],ptrAtom_Path[intTop_Node],intArray_Length);

				//Put attached node ptrTemp_Bond->ptrNextBond onto the back of the qeue.
				ptrTemp2_Que->Add_Que(intTop_Node,ptrTemp_Bond->intAttachedAtom);
				ptrTemp2_Que=ptrTemp2_Que->ptrNextQue;

			}
		}

		ptrTemp_Bond=ptrTemp_Bond->ptrNextBond;
	}

	//Pop off the top pointer.
	if(ptrTemp1_Que) {
		ptrTemp3_Que=ptrTemp1_Que->ptrNextQue;
		ptrTemp1_Que->ptrNextQue=0;
		delete ptrTemp1_Que;
		ptrTemp1_Que=ptrTemp3_Que;
	}
}
//Clean up dynamic memory allocations.
for(j=0;j<intNumberAtoms;j++) {
	delete[] ptrAtom_Path[j];
}
delete[] ptrAtom_Path;
if(ptrTemp1_Que) delete ptrTemp1_Que;
return 0;

}

//Routine used in Figueras_SSSR() to trim a node from the query structure for further processing.
inline void
ChemSeqID::Trim_SSSR(unsigned _int32 intTrim_Loc,unsigned _int8 *ptrDegree,unsigned _int32 *ptrTrim_Set) {

const unsigned _int8 intBit_Size=32;	//Bit size of integer used in Trim_Set[] array.
unsigned _int32 intBit_Loc;				//Bit location in array integer of current node.
unsigned _int32 intArray_Loc;			//Array location in array of current node.
AtomBond *ptrTemp_Bond;					//Bond pointer used to loop through neighbors of current node.

//Set "trimmed node" bit in trimSet.
intArray_Loc=intTrim_Loc/intBit_Size;
intBit_Loc=intTrim_Loc%intBit_Size;
ptrTrim_Set[intArray_Loc] |= (1<<intBit_Loc);

if(ptrDegree[intTrim_Loc]) {

	//Set "trimmed node" degree to zero.
	ptrDegree[intTrim_Loc]=0;

	//Decrement the degree of "trimmed node's" neighbors.
	ptrTemp_Bond=ptrMolecule->ptrAtom[intTrim_Loc].ptrNextBond;
	while(ptrTemp_Bond) {
		if(ptrDegree[ptrTemp_Bond->intAttachedAtom])ptrDegree[ptrTemp_Bond->intAttachedAtom]--;
		ptrTemp_Bond=ptrTemp_Bond->ptrNextBond;
	}
}

}

//Ring perception routine based on Figueras (1996) SSSR ring perception algorithm. Returns the SSSR in the ChemSeqID.
unsigned _int8
ChemSeqID::Figueras_SSSR() {

unsigned _int8 *ptrDegree=0;	//Array stores current degree of each atom during processing.
unsigned _int32 *ptrNodes_N2=0;	//Array stores current atom locations of degree 2.
unsigned _int32 *ptrTrim_Set=0;	//Array stores the "trimmed" nodes as bits in the trim set.
unsigned _int32 *ptrFull_Set=0;	//Array stores the entire molecular structure as set bits in the full set.
unsigned _int32 *ptrRing_Set=0;	//Array stores ring detections as set bits in an array representing the query structure.
unsigned _int32 i;				//Junk variable.
unsigned _int32 intN2_Cntr;		//Counter denoting current number of N2 atoms.
unsigned _int32 intLast_Change;	//Stores location of last "trimmed" node in the structure.
unsigned _int32 intMin_Place;	//Stores location of node with minimum degree in structure.
unsigned _int32 intArray_Length;//Size of query structure array in bits.
unsigned _int32 intRing_Size=0;	//Size of current detected ring.
unsigned _int32 intRing_Type=0;	//Type of detected ring (aryl, alicyclic, etc.).
const unsigned _int32 intZero=0;//32 bit integer set to zero
const unsigned _int8 intBit_Size=32;	//Size of integer used in storage arrays in bits.

SSSR *ptrTemp_SSSR;				//Temporary que node pointer.

//Fill SSSR check flag.
Flags.SSSR_Fill=1;

//Create and initialize arrays.
intArray_Length=(intNumberAtoms-1)/intBit_Size+1;
ptrFull_Set=new unsigned _int32[intArray_Length];
ptrTrim_Set=new unsigned _int32[intArray_Length];
ptrRing_Set=new unsigned _int32[intArray_Length];
ptrDegree=new unsigned _int8[intNumberAtoms];
ptrNodes_N2=new unsigned _int32[intNumberAtoms];

for(i=0;i<intArray_Length;i++) {
	ptrTrim_Set[i]=0;
	if(i != (intArray_Length-1))ptrFull_Set[i]=~intZero;
	else if(intNumberAtoms == intBit_Size) ptrFull_Set[i]=~intZero;
	else ptrFull_Set[i]=((1<<(intNumberAtoms%intBit_Size))-1);
}
for(i=0;i<intNumberAtoms;i++) {
	ptrDegree[i]=ptrMolecule->ptrAtom[i].Atom_ID.Num_NH_Cnt;
	ptrNodes_N2[i]=0;
	//If node is of degree zero then trim from structure (free ions, etc.)
	if(!ptrDegree[i]) Trim_SSSR(i,ptrDegree,ptrTrim_Set);
}

//Look for rings until entire structure is searched.
while(!Compare_Set(ptrFull_Set,ptrTrim_Set,intArray_Length)) {

	//Initialize loop variables.
	i=0;
	intLast_Change=0;

	//Prune all non-cyclic entities from structure.
	do{

		if(ptrDegree[i]<2) {
			if(ptrDegree[i]) intLast_Change=i;
			Trim_SSSR(i,ptrDegree,ptrTrim_Set);
		}
		if(i<(intNumberAtoms-1)) i++;
		else i=0;
	
	}while(i != intLast_Change);
	
	//Determine if structure contains a non-null node.
	intMin_Place=0;
	for(i=0;i<intNumberAtoms;i++) {
		if(ptrDegree[i]) {
			intMin_Place=i;
			break;
		}
	}
	//If no non-null node exists, exit loop.
	if(i==intNumberAtoms) break;

	//Find location of minimum degree.
	for(i=0;i<intNumberAtoms;i++) {
		if(ptrDegree[i] && ptrDegree[i]<ptrDegree[intMin_Place]) intMin_Place=i;
	}

	intN2_Cntr=0;
	//Add nodes of degree N2 in fullSet-trimSet to nodes N2.
	for(i=0;i<intNumberAtoms;i++) {
		if(ptrDegree[i]==2) {
			ptrNodes_N2[intN2_Cntr]=i;
			intN2_Cntr++;
		}
	}
	
	//BFS process for minimum nodes of degree two.
	if(ptrDegree[intMin_Place]==2) {

		//Retrieve ring if present. //If ring detection occurred, then add ring to SSSR.
		for(i=0;i<intN2_Cntr;i++) {

			//Get ring routine.
			if(Get_Ring(ptrNodes_N2[i],ptrRing_Set,ptrDegree,intRing_Size,intArray_Length)) {
			
				if(!ptrSSSR) {
					ptrSSSR=new SSSR(intRing_Type,intRing_Size,ptrRing_Set,intArray_Length);
					ptrTemp_SSSR=ptrSSSR;
				}
				else if(Check_Duplicate(ptrRing_Set,intArray_Length)) {
					ptrTemp_SSSR->ptrNextSSSR=new SSSR(intRing_Type,intRing_Size,ptrRing_Set,intArray_Length);
					ptrTemp_SSSR=ptrTemp_SSSR->ptrNextSSSR;
				}

				//Determine the ring type of detected ring.
				Calc_Ring_Type(ptrTemp_SSSR);
			}
		}
		//Trim (remove from search) the investigated nodes of degree 2.
		for(i=0;i<intN2_Cntr;i++) {
			Trim_SSSR(ptrNodes_N2[i],ptrDegree,ptrTrim_Set);
		}
	}
	//BFS process for minimum nodes of degree three.
	else if(ptrDegree[intMin_Place]==3) {

		//Get ring routine.
		if(Get_Ring(intMin_Place,ptrRing_Set,ptrDegree,intRing_Size,intArray_Length)) {

			if(!ptrSSSR) {
				ptrSSSR=new SSSR(intRing_Type,intRing_Size,ptrRing_Set,intArray_Length);
				ptrTemp_SSSR=ptrSSSR;
			}
			else if(Check_Duplicate(ptrRing_Set,intArray_Length)) {
				ptrTemp_SSSR->ptrNextSSSR=new SSSR(intRing_Type,intRing_Size,ptrRing_Set,intArray_Length);
				ptrTemp_SSSR=ptrTemp_SSSR->ptrNextSSSR;
			}

			//Determine the ring type of detected ring.
			Calc_Ring_Type(ptrTemp_SSSR);

			//Call routine to determine all rings attached to the node elimination candidate and then eliminate the node.
			Check_Nodes(intMin_Place,intRing_Size,ptrRing_Set,ptrTrim_Set,ptrDegree,intArray_Length);

		}

	}

}

//Clean up dynamically allocated memory.
if(ptrRing_Set) delete[] ptrRing_Set;
if(ptrDegree) delete[] ptrDegree;
if(ptrNodes_N2) delete[] ptrNodes_N2;
if(ptrFull_Set) delete[] ptrFull_Set;
if(ptrTrim_Set) delete[] ptrTrim_Set;

//Return values.
if(ptrSSSR) return 1;
else return 0;

}

//This routine determines whether the bond oriented molecular feature is present and in what quantities in the query molecule.
unsigned _int16
ChemSeqID::Retrieve_Bond_MF(ChemSeqID *ptrBetaIDNode) {

	unsigned _int16 intDetect_Quant=0;		//Counter variable for number of MF detections.
	unsigned _int32 i;						//Junk counter variable.
	AtomBond *ptrTemp_Bond;					//Temporary bond pointer used to loop through each atoms connections (augmented atom complex).
	

	//Check if perception is necessary (i.e., bond must either be or not be in a ring). Fill Search ID ring search codes.
	if(!Flags.RSC_Fill && (ptrBetaIDNode->MF_ID.Bond_Constraint&(1<<0))) {
		if(ptrBetaIDNode->MF_ID.Bond_Constraint !=3) Calc_RSC();
	}

	//Loop through the query graph .
	for(i=0;i<intNumberAtoms;i++) {

		ptrTemp_Bond=(ptrMolecule->ptrAtom[i]).ptrNextBond;

		while(ptrTemp_Bond) {

			//Prevent double counting of bonds due to previous atoms.
			if(ptrTemp_Bond->intAttachedAtom>=i) {

				//If MF's bond type is set to all bonds, the bonds equal or pass the greater than or less than comparison.
				if(ptrTemp_Bond->Bond_ID.Bond_Type==7 || ptrTemp_Bond->Bond_ID.Bond_Type==ptrBetaIDNode->MF_ID.Bond_Type || (ptrBetaIDNode->MF_ID.B_Greater_Than && ptrBetaIDNode->MF_ID.Bond_Type>ptrTemp_Bond->Bond_ID.Bond_Type) || (ptrBetaIDNode->MF_ID.B_Less_Than && ptrBetaIDNode->MF_ID.Bond_Type<ptrTemp_Bond->Bond_ID.Bond_Type)) {

					//If "not in ring" bond constraints correspond.
					if((ptrBetaIDNode->MF_ID.Bond_Constraint & (1<<0)) && (ptrMolecule->ptrAtom[i].Search_ID.Not_Ring || ptrMolecule->ptrAtom[ptrTemp_Bond->intAttachedAtom].Search_ID.Not_Ring)) {

						//If the not bond flag is not set.
						if(!ptrBetaIDNode->MF_ID.Not_Bond) intDetect_Quant++;
					}
					//Else if "in ring" bond constraints correspond.
					else if((ptrBetaIDNode->MF_ID.Bond_Constraint & (1<<1)) && (ptrMolecule->ptrAtom[i].Search_ID.In_Ring && ptrMolecule->ptrAtom[ptrTemp_Bond->intAttachedAtom].Search_ID.In_Ring)) {

						//If the not bond flag is not set.
						if(!ptrBetaIDNode->MF_ID.Not_Bond) intDetect_Quant++;
					}

				}
				//Else if the bonds do not equal or do not pass the greater than or less than comparison.
				else {

					//If "not in ring" bond constraints correspond.
					if((ptrBetaIDNode->MF_ID.Bond_Constraint & (1<<0)) && (ptrMolecule->ptrAtom[i].Search_ID.Not_Ring || ptrMolecule->ptrAtom[ptrTemp_Bond->intAttachedAtom].Search_ID.Not_Ring)) {

						//If the not bond flag is set.
						if(ptrBetaIDNode->MF_ID.Not_Bond) intDetect_Quant++;
					}
					//Else if "in ring" bond constraints correspond.
					else if((ptrBetaIDNode->MF_ID.Bond_Constraint & (1<<1)) && (ptrMolecule->ptrAtom[i].Search_ID.In_Ring && ptrMolecule->ptrAtom[ptrTemp_Bond->intAttachedAtom].Search_ID.In_Ring)) {

						//If the not bond flag is not set.
						if(ptrBetaIDNode->MF_ID.Not_Bond) intDetect_Quant++;
					}
				}
		
				ptrTemp_Bond=ptrTemp_Bond->ptrNextBond;
			}
		}
	}

return intDetect_Quant;
}

//This routine determines whether the ring oriented molecular feature is present and in what quantities in the query molecule.
unsigned _int16
ChemSeqID::Retrieve_Ring_MF(ChemSeqID *ptrBetaIDNode) {

	unsigned _int16 intDetect_Quant=0;		//Counter variable for number of MF detections.
	SSSR *ptrTemp_SSSR;						//Temporary SSSR pointer
	
	//Check if perception is necessary (i.e., Already processed?)
	if(!Flags.SSSR_Fill) {
		Figueras_SSSR();
	}

	ptrTemp_SSSR=ptrSSSR;

	//Loop through "smallest set of smallest ring" list to determine how many MF matches exist.
	while(ptrTemp_SSSR) {

		//If ring size and ring type correspond or ring type and greater than less than comparisons correspond.
		if((ptrTemp_SSSR->Ring_Type & ptrBetaIDNode->MF_ID.Ring_Type) && (ptrTemp_SSSR->Num_Members==ptrBetaIDNode->MF_ID.Ring_Size || (ptrBetaIDNode->MF_ID.R_Greater_Than && ptrTemp_SSSR->Num_Members>ptrBetaIDNode->MF_ID.Ring_Size) || (ptrBetaIDNode->MF_ID.R_Less_Than && ptrTemp_SSSR->Num_Members<ptrBetaIDNode->MF_ID.Ring_Size))) {

			//If not ring flag is not set, increment detection.
			if(!ptrBetaIDNode->MF_ID.Not_Ring) intDetect_Quant++;
		}
		//If ring size and ring type don't correspond or ring type and greater than less than comparisons don't correspond.
		else {
			if(ptrBetaIDNode->MF_ID.Not_Ring) intDetect_Quant++;
		}

		ptrTemp_SSSR=ptrTemp_SSSR->ptrNextSSSR;
	}

	return intDetect_Quant;
}

//This routine fills molecular feature bitfield if the string is a Molecular Feature or exits as a failure if the << >> string is a MF constraining operator.
void
ChemSeqID::Fill_MF() {

unsigned _int8 intFeatureFlag=0;        //  ...| aryl ring | alicyclic ring | ring bond | non-ring bond |   
unsigned _int32 intTempCount=0;			//Temporary counter variable
unsigned _int32 intNum_Chars;			//Stores number of characters in string

//Check for feature commands located inside ' <<___>> '
intFeatureFlag=0;
intNum_Chars=(strlen(chrChemSyntaxString)-1);

if(chrChemSyntaxString[intTempCount]=='<' && chrChemSyntaxString[intTempCount+1] =='<') { 
	//Check whether subfragment feature operator (<<) initiates a molecular feature string or a
    //molecular feature operator for a substructure. 
    if(!intTempCount) intTempCount =2;
	else return;

    while(!(chrChemSyntaxString[intTempCount-1] == '>' && chrChemSyntaxString[intTempCount] == '>')) {
		intTempCount++;
    }
  
	//Molecular feature grouping (<<___>>) is a molecular feature string. No subfragment string follows the (<__>) operator.
    if(intTempCount == intNum_Chars) {

		intTempCount=2;
		//Set bit denoting that the molecular feature bitfield has been filled.
		Flags.MF_Fill=1;
		//Set bit denoting that the quantity key bitfield has been filled.
		Flags.QK_Fill=1;

		while(intTempCount < intNum_Chars-1){
			switch(chrChemSyntaxString[intTempCount]) {
            //Not operator (essentially serves as an unspecifier at this point)
            case '!':
				if(chrChemSyntaxString[intTempCount+1]=='R' || chrChemSyntaxString[intTempCount+1]=='r') MF_ID.Not_Ring=1;
                else if(chrChemSyntaxString[intTempCount+1]=='B' || chrChemSyntaxString[intTempCount+1]=='b') MF_ID.Not_Bond=1;
                break;
            //Flag for # of alicyclic rings
            case 'R':
				MF_ID.Ring_Feature=1;
				MF_ID.Ring_Type |= (1<<0);   
                if(!intQuantKey.Num_Rings && !MF_ID.Not_Ring) intQuantKey.Num_Rings++;
				intFeatureFlag |= (1<<2);
                break;
            //Flag for # of aryl rings
            case 'r':
                MF_ID.Ring_Feature=1;
				MF_ID.Ring_Type |= (1<<1);
                if(!intQuantKey.Num_Rings && !MF_ID.Not_Ring) intQuantKey.Num_Rings++;
				intFeatureFlag |= (1<<3);
                break;
            //Flag for # of non-ring bonds
            case 'B':
                MF_ID.Bond_Feature=1;
				intFeatureFlag |= (1<<0);
                break;
            //Flag for # of ring bonds
            case 'b':
				MF_ID.Bond_Feature=1;
				intFeatureFlag |= (1<<1);
                break;
            //Less than symbol
            case '<':
                if(intFeatureFlag & (3<<2)) MF_ID.R_Less_Than=1;
                else if(intFeatureFlag & (3<<0)) MF_ID.B_Less_Than=1;
                break;
            //Greater than symbol
            case '>':
                if(intFeatureFlag & (3<<2)) MF_ID.R_Greater_Than=1;
                else if(intFeatureFlag & (3<<0)) MF_ID.B_Greater_Than=1;
                break;
            //Numbers used to specify bonds and ring sizes
            default:
				if(chrChemSyntaxString[intTempCount] >47 && chrChemSyntaxString[intTempCount] <58) {
                                   
					switch(intFeatureFlag) {
					//Ring and non-ring bonds
					case 3:
                                        
					//Ring bond
					case 2:
						
						MF_ID.Bond_Constraint |= intFeatureFlag;	

						switch(chrChemSyntaxString[intTempCount]-48) {
						//All bonds
						case 0:
							MF_ID.Bond_Type=7;
							break;
						//Single bond
						case 1:
							MF_ID.Bond_Type=1;
							break;
						//Double bond
						case 2:     
							MF_ID.Bond_Type=2;
							if(!MF_ID.Not_Bond) intQuantKey.Num_DB++;
							break;
						//Triple bond
                        case 3:
                            MF_ID.Bond_Type=3;
                            if(!MF_ID.Not_Bond) intQuantKey.Num_TB++;
							break;
                        //Aryl (delocalized) bond
                        case 4:
                            MF_ID.Bond_Type=4;
                            break;
                        }
                        if(intFeatureFlag<3)break;
                    //Non-ring bond
                    case 1:

						MF_ID.Bond_Constraint |= intFeatureFlag;	

                        switch(chrChemSyntaxString[intTempCount]-48) {
						//All bonds
						case 0:
							MF_ID.Bond_Type=7;
							break;
						//Single bond
                        case 1:
                            MF_ID.Bond_Type=1;
                            break;
                        //Double bond
                        case 2:     
                            MF_ID.Bond_Type=2;
                            break;
                        //Triple bond
                        case 3:
                            MF_ID.Bond_Type=3;
                            break;
                        }
                        break;
                    //Aryl and Alicyclic rings
                    case 12:
                    //Aryl ring
                    case 8:
                    //Alicyclic ring
                    case 4:
                        MF_ID.Ring_Size = (chrChemSyntaxString[intTempCount]-48);
                        break;
                    }
                }
            }               
            intTempCount++;
        }
    }
}

}

//Scan function performs a simple atom quantity "screen ID filling" of the atom quantity key bitfield for the 
//molecule to use in an atom quantity screening to determine whether it is worthwhile to parse into an ADT
//(Currently only supports =O for quantity keys greater than one character.)
void
ChemSeqID::SMILES_QK_Fill() {

     unsigned _int32 intCharCount=0;         //Number of characters variable
     unsigned _int32 intCharTempCnt;         //Temporary character counter variable
	 unsigned _int32 intAttach_Cnt=0;		 //Nested branches counter
     unsigned _int16 intTempMolFeature=0;    //Temporary molecular feature bit field
	 unsigned _int16 intRing_Number=0;		 //Stores ring number retrieved from SMILES string
     unsigned _int8 intCurrentBond=1;        //Variable denoting current bond in quantity fill process
     unsigned _int8 intBracketFlag=0;        //Flag determining if current atom is a member of an atom bracket [] group
     unsigned _int16 intRingArray[60]={0};   //Array used to store ring numbers
     unsigned _int8 intComboFlag;            //Flag to determine if a combo, [X,x], was detected.
     unsigned _int8 intOF_Cntr=0;            //Molecular operation flag counter
     unsigned _int8 intOF_BF=0;              //Flag bit field for molecular operators  ...| 3~MF operator | 2~Not atom | 1~Dechar atom |0~blank | 
     unsigned _int8 intGF_BF=0;              //Flag bit filed for grouping flags
     unsigned _int8 intOF_Que[5]={0};        //Que array storing type of despecification operator at appropriate level of occurrence
     unsigned _int8 intPast_Atom[100]={0};   //Array storing atom type of previous attachments.
     
     intNumberAtoms=0;						 //Number of atoms variable

     //Set intFillBit to show that atom quantity screen bit has been processed.
     Flags.QK_Fill=1;

     while(chrChemSyntaxString[intCharCount] != '\0') {

		  //If character is despecified by either the not or decharacterization operators, then don't enter QK incrementing block.
          if(!(intOF_BF & (1<<1)) && !(intOF_BF & (1<<2))) {

               switch(chrChemSyntaxString[intCharCount]) {

               //Elements (atoms) beginning with 'H'
               case 'H':
				   if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
						//Helium
						case 'e':
							intPast_Atom[intAttach_Cnt]=2;
							intQuantKey.Num_Etc++;
							break;
						//Hafnium
						case 'f':
							intPast_Atom[intAttach_Cnt]=72;
							intQuantKey.Num_Etc++;
							break;
	                    //Mercury
		                case 'g':
							intPast_Atom[intAttach_Cnt]=80;
							intQuantKey.Num_Etc++;
							break;
			            //Holmium
				        case 'o':
							intPast_Atom[intAttach_Cnt]=67;
							intQuantKey.Num_Etc++;
							break;
						//Hydrogen
				        default:
							break;
						}
				   }
				   if(intBracketFlag) intBracketFlag++;
				   break;
               //Elements beginning with 'L'
               case 'L':
				   if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
		                //Lithium
			            case 'i':
							intPast_Atom[intAttach_Cnt]=3;
							intQuantKey.Num_Etc++;
							break;
				        //Lanthanum
					    case 'a':
							intPast_Atom[intAttach_Cnt]=57;
							intQuantKey.Num_Etc++;
							break;
						//Lutetium
	                    case 'u':
							intPast_Atom[intAttach_Cnt]=71;
							intQuantKey.Num_Etc++;
							break;
					    }
				   }
				   if(intBracketFlag) intBracketFlag++;
		           break;
               //Elements beginning with 'B'
               case 'B':
				   if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
		                //Bromine
			            case 'r':
							intPast_Atom[intAttach_Cnt]=35;
							intQuantKey.Num_Br++;
							break;
			            //Beryllium
		                case 'e':
							intPast_Atom[intAttach_Cnt]=4;
							intQuantKey.Num_Etc++;
							break;
				        //Barium
					    case 'a':
							intPast_Atom[intAttach_Cnt]=56;
							intQuantKey.Num_Etc++;
							break;
						//Bismuth
	                    case 'i':
							intPast_Atom[intAttach_Cnt]=83;
							intQuantKey.Num_Etc++;
							break;
		                //Berkelium
			            case 'k':
							intPast_Atom[intAttach_Cnt]=97;
							intQuantKey.Num_Etc++;
							break;
					    //Boron
	                    default:
							intPast_Atom[intAttach_Cnt]=5;
							intQuantKey.Num_Etc++;
							break;
						}
				   }
				   if(intBracketFlag<2) intBracketFlag++;
				   break;
               //Elements beginning with 'C'
               case 'C':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
	                    //Chlorine
		                case 'l':
			            	intPast_Atom[intAttach_Cnt]=17;
							intQuantKey.Num_Cl++;
							break;
	                    //Calcium
		                case 'a':
							intPast_Atom[intAttach_Cnt]=20;
							intQuantKey.Num_Etc++;
							break;
			            //Chromium
				        case 'r':
							intPast_Atom[intAttach_Cnt]=24;
							intQuantKey.Num_Etc++;
							break;
					    //Cobalt
						case 'o':
							intPast_Atom[intAttach_Cnt]=27;
							intQuantKey.Num_Etc++;
							break;
	                    //Copper
		                case 'u':
							intPast_Atom[intAttach_Cnt]=29;
							intQuantKey.Num_Etc++;
							break;
			            //Cadmium
				        case 'd':
							intPast_Atom[intAttach_Cnt]=48;
							intQuantKey.Num_Etc++;
							break;
					    //Cesium
						case 's':
							intPast_Atom[intAttach_Cnt]=55;
							intQuantKey.Num_Etc++;
							break;
	                    //Cerium
		                case 'e':
							intPast_Atom[intAttach_Cnt]=58;
							intQuantKey.Num_Etc++;
							break;
			            //Curium
				        case 'm':
					        intPast_Atom[intAttach_Cnt]=96;
							intQuantKey.Num_Etc++;
							break; 
			            //Carbon
				        default:
					         
							//Increment the carbon double bonded to carbon key.
							if(intPast_Atom[intAttach_Cnt] == 6 && intCurrentBond==2) intQuantKey.Num_CDBC++;
							//Increment the carbon double bonded to oxygen key.
							else if(intPast_Atom[intAttach_Cnt] == 8 && intCurrentBond==2) intQuantKey.Num_CDBO++;
							//Increment the carbon triple bonded to nitrogen key.
							else if(intPast_Atom[intAttach_Cnt]== 7 && intCurrentBond==3) intQuantKey.Num_CTBN++;
	
							intPast_Atom[intAttach_Cnt]=6;
							intQuantKey.Num_C++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
				    break;
               //Elements beginning with 'N'
               case 'N':
				    if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
		                //Neon
			            case 'e':
							intPast_Atom[intAttach_Cnt]=10;
							intQuantKey.Num_Etc++;
							break;
				        //Sodium
					    case 'a':
							intPast_Atom[intAttach_Cnt]=11;
							intQuantKey.Num_Etc++;
							break;
						//Nickel
	                    case 'i':
							intPast_Atom[intAttach_Cnt]=28;
							intQuantKey.Num_Etc++;
							break;
		                //Niobium
			            case 'b':
							intPast_Atom[intAttach_Cnt]=41;
							intQuantKey.Num_Etc++;
							break;
				        //Neodymium
					    case 'd':
							intPast_Atom[intAttach_Cnt]=60;
							intQuantKey.Num_Etc++;
							break;
						//Neptunium
	                    case 'p':
							intPast_Atom[intAttach_Cnt]=93;
							intQuantKey.Num_Etc++;
							break;
		                //Nobellium
			            case 'o':
				            intPast_Atom[intAttach_Cnt]=102;
							intQuantKey.Num_Etc++;
							break; 
		                //Nitrogen
			            default:
							//Increment the carbon triple bonded to nitrogen key.
							if(intPast_Atom[intAttach_Cnt] == 6 && intCurrentBond==3) intQuantKey.Num_CTBN++;
								 
							intPast_Atom[intAttach_Cnt]=7;
							intQuantKey.Num_N++;
							break;
						}
					}
					if(intBracketFlag) intBracketFlag++;
					break;
               //Element beginning with 'O'
               case 'O':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
		                //Osmium
			            case 's':
							intPast_Atom[intAttach_Cnt]=76;
							intQuantKey.Num_Etc++;
							break;
		                //Oxygen
			            default:
				            //Increment carbon double bonded to oxygen key.
							if(intPast_Atom[intAttach_Cnt] == 6 && intCurrentBond==2) intQuantKey.Num_CDBO++;
							 							 
							intPast_Atom[intAttach_Cnt]=8;
							intQuantKey.Num_O++;
							break;
					    }
					}
					if(intBracketFlag) intBracketFlag++;
					break;
               //Elements beginning with 'F'
               case 'F':    
				   if(intBracketFlag<2) { 
					   switch(chrChemSyntaxString[intCharCount+1]) {
		                //Iron
			            case 'e':
							intPast_Atom[intAttach_Cnt]=26;
							intQuantKey.Num_Etc++;
							break;
				        //Francium
					    case 'r':
							intPast_Atom[intAttach_Cnt]=87;
							intQuantKey.Num_Etc++;
							break;
						//Fermium
	                    case 'm':
							intPast_Atom[intAttach_Cnt]=100;
							intQuantKey.Num_Etc++;
							break;
	                    //Fluorine
		                default:
							intPast_Atom[intAttach_Cnt]=9;
							intQuantKey.Num_F++;
							break;
					   }
				   }
				   if(intBracketFlag) intBracketFlag++;
				   break;
               //Elements beginning with 'M'
               case 'M':
				   if(intBracketFlag<2) {
	                    switch(chrChemSyntaxString[intCharCount+1]) {
		                //Magnesium
			            case 'g':
							intPast_Atom[intAttach_Cnt]=12;
							intQuantKey.Num_Etc++;
							break;
				        //Manganese
					    case 'n':
							intPast_Atom[intAttach_Cnt]=25;
							intQuantKey.Num_Etc++;
							break;
						//Molybdenum
	                    case 'o':
							intPast_Atom[intAttach_Cnt]=42;
							intQuantKey.Num_Etc++;
							break;
		                //Mendelevium
			            case 'd':
				            intPast_Atom[intAttach_Cnt]=101;
							intQuantKey.Num_Etc++;
							break; 
		                }
				   }
				   if(intBracketFlag) intBracketFlag++;
			       break;
               //Elements beginning with 'A'
               case 'A':
					if(intBracketFlag<2) {
					   switch(chrChemSyntaxString[intCharCount+1]) {
	                    //Aluminum
		                case 'l':
							intPast_Atom[intAttach_Cnt]=13;
							intQuantKey.Num_Etc++;
							break;
			            //Argon
				        case 'r':
							intPast_Atom[intAttach_Cnt]=18;
							intQuantKey.Num_Etc++;
							break;
					    //Arsenic
						case 's':
							intPast_Atom[intAttach_Cnt]=33;
							intQuantKey.Num_Etc++;
							break;
	                    //Silver
		                case 'g':
							intPast_Atom[intAttach_Cnt]=47;
							intQuantKey.Num_Etc++;
							break;
			            //Gold
				        case 'u':
							intPast_Atom[intAttach_Cnt]=79;
							intQuantKey.Num_Etc++;
							break;
					    //Astatine
						case 't':
							intPast_Atom[intAttach_Cnt]=85;
							intQuantKey.Num_Etc++;
							break;
	                    //Actinum
		                case 'c':
							intPast_Atom[intAttach_Cnt]=89;
							intQuantKey.Num_Etc++;
							break;
			            //Americum
				        case 'm':
							intPast_Atom[intAttach_Cnt]=95;
							intQuantKey.Num_Etc++;
							break;
						//Aliphatic wild card atom
				        default:
					         break;
						}
					}
					if(intBracketFlag) intBracketFlag++;
					break;
               //Elements beginning with 'S'
               case 'S':
				   if(intBracketFlag<2) {
	                    switch(chrChemSyntaxString[intCharCount+1]) {
		                //Silicon
			            case 'i':
							intPast_Atom[intAttach_Cnt]=14;
							intQuantKey.Num_Etc++;
							break;
				        //Scandium
					    case 'c':
							intPast_Atom[intAttach_Cnt]=21;
							intQuantKey.Num_Etc++;
							break;
						//Selenium
	                    case 'e':
							intPast_Atom[intAttach_Cnt]=34;
							intQuantKey.Num_Etc++;
							break;
		                //Strontium
			            case 'r':
							intPast_Atom[intAttach_Cnt]=38;
							intQuantKey.Num_Etc++;
							break;
				        //Tin
					    case 'n':
							intPast_Atom[intAttach_Cnt]=50;
							intQuantKey.Num_Etc++;
							break;
						//Antimony
	                    case 'b':
							intPast_Atom[intAttach_Cnt]=51;
							intQuantKey.Num_Etc++;
							break;
		                //Samarium
			            case 'm':
							intPast_Atom[intAttach_Cnt]=62;
							intQuantKey.Num_Etc++;
							break;
		                //Sulfur
			            default:
							intPast_Atom[intAttach_Cnt]=16;
							intQuantKey.Num_S++;
							break;
		                }
				   }
				   if(intBracketFlag) intBracketFlag++;
			       break;
               //Elements beginning with 'P'
               case 'P':
					if(intBracketFlag<2) {
	                    switch(chrChemSyntaxString[intCharCount+1]) {
		                //Paladium
			            case 'd':
							intPast_Atom[intAttach_Cnt]=46;
							intQuantKey.Num_Etc++;
							break;
				        //Platinum
					    case 't':
							intPast_Atom[intAttach_Cnt]=78;
							intQuantKey.Num_Etc++;
							break;
						//Lead
	                    case 'b':
							intPast_Atom[intAttach_Cnt]=82;
							intQuantKey.Num_Etc++;
							break;
		                //Polonium
			            case 'o':
							intPast_Atom[intAttach_Cnt]=84;
							intQuantKey.Num_Etc++;
							break;
				        //Praseodymium
					    case 'r':
							intPast_Atom[intAttach_Cnt]=59;
							intQuantKey.Num_Etc++;
							break;
						//Promethium
	                    case 'm':
							intPast_Atom[intAttach_Cnt]=61;
							intQuantKey.Num_Etc++;
							break;
		                //Protactinum
			            case 'a':
							intPast_Atom[intAttach_Cnt]=91;
							intQuantKey.Num_Etc++;
							break;
				        //Plutonium
					    case 'u':
						    intPast_Atom[intAttach_Cnt]=94;
							intQuantKey.Num_Etc++;
							break;
			            //Phosphorous
				        default:
					        intPast_Atom[intAttach_Cnt]=15;
							intQuantKey.Num_P++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
				    break;
               //Elements beginning with 'K'
               case 'K':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
	                    //Krypton
		                case 'r':
							intPast_Atom[intAttach_Cnt]=36;
							intQuantKey.Num_Etc++;
							break;
				        //Potassium
					    default:
							intPast_Atom[intAttach_Cnt]=19;
							intQuantKey.Num_Etc++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
				    break;
               //Elements beginning with 'T'
               case 'T':
					if(intBracketFlag<2) {
	                    switch(chrChemSyntaxString[intCharCount+1]) {
		                //Titanium
			            case 'i':
							intPast_Atom[intAttach_Cnt]=22;
							intQuantKey.Num_Etc++;
							break;
				        //Technetium
						case 'c':
							intPast_Atom[intAttach_Cnt]=43;
							intQuantKey.Num_Etc++;
							break;
						//Tellurium
	                    case 'e':
							intPast_Atom[intAttach_Cnt]=52;
							intQuantKey.Num_Etc++;
							break;
		                //Tantalum
			            case 'a':
							intPast_Atom[intAttach_Cnt]=73;
							intQuantKey.Num_Etc++;
							break;
				        //Thallium
					    case 'l':
							intPast_Atom[intAttach_Cnt]=81;
							intQuantKey.Num_Etc++;
							break;
						//Terbium
	                    case 'b':
							intPast_Atom[intAttach_Cnt]=65;
							intQuantKey.Num_Etc++;
							break;
		                //Thorium
			            case 'h':
							intPast_Atom[intAttach_Cnt]=90;
							intQuantKey.Num_Etc++;
							break;
				        //Thulium
					    case 'm':
							intPast_Atom[intAttach_Cnt]=69;
							intQuantKey.Num_Etc++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
					break;
               //Elements beginning with 'V'
               //Vanadium
               case 'V':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=23;
						intQuantKey.Num_Etc++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Elements beginning with 'Z'
               case 'Z':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
	                    //Zinc
		                case 'n':
							intPast_Atom[intAttach_Cnt]=30;
							intQuantKey.Num_Etc++;
							break;
			            //Zirconium
				        case 'r':
							intPast_Atom[intAttach_Cnt]=40;
							intQuantKey.Num_Etc++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
				    break;
               //Elements beginning with 'G' 
               case 'G':     
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
						//Gallium
						case 'a':
							intPast_Atom[intAttach_Cnt]=31;
							intQuantKey.Num_Etc++;
							break;
	                    //Germanium
		                case 'e':
							intPast_Atom[intAttach_Cnt]=32;
							intQuantKey.Num_Etc++;
							break;
			            //Gadolinium
				        case 'd':
							intPast_Atom[intAttach_Cnt]=64;
							intQuantKey.Num_Etc++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
				    break;
               //Elements beginning with 'R'
               case 'R':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
	                    //Rubidium
		                case 'b':
							intPast_Atom[intAttach_Cnt]=37;
							intQuantKey.Num_Etc++;
							break;
			            //Ruthenium
				        case 'u':
							intPast_Atom[intAttach_Cnt]=44;
							intQuantKey.Num_Etc++;
							break;
					    //Rhodium
						case 'h':
							intPast_Atom[intAttach_Cnt]=45;
							intQuantKey.Num_Etc++;
							break;
	                    //Rhenium
		                case 'e':
							intPast_Atom[intAttach_Cnt]=75;
							intQuantKey.Num_Etc++;
							break;
			            //Radon
				        case 'n':
							intPast_Atom[intAttach_Cnt]=86;
							intQuantKey.Num_Etc++;
							break;
					    //Radium
						case 'a':
							intPast_Atom[intAttach_Cnt]=88;
							intQuantKey.Num_Etc++;
							break;
				        }
					}
					if(intBracketFlag) intBracketFlag++;
					break;
               //Elements beginning with 'Y'
               case 'Y':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]){
	                    //Ytterbium
		                case 'b':
							intPast_Atom[intAttach_Cnt]=70;
							intQuantKey.Num_Etc++;
							break;
			            //Yttrium
				        default:
							intPast_Atom[intAttach_Cnt]=39;
							intQuantKey.Num_Etc++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
				    break;
               //Elements beginning with 'I'
               case 'I':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
	                    //Indium
		                case 'n':
							intPast_Atom[intAttach_Cnt]=49;
							intQuantKey.Num_Etc++;
							break;
			            //Iridium
				        case 'r':
							intPast_Atom[intAttach_Cnt]=77;
							intQuantKey.Num_Etc++;
							break;
			            //Iodine
				        default:
							intPast_Atom[intAttach_Cnt]=53;
							intQuantKey.Num_I++;
							break;
			            }
					}
					if(intBracketFlag) intBracketFlag++;
				    break;
               //Elements beginning with 'X'
               //Xenon
               case 'X':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=54;
						intQuantKey.Num_Etc++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Elements beginning with 'W'
               //Wolfram
               case 'W':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=74;
						intQuantKey.Num_Etc++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Elements beginning with 'E'
               case 'E':
					if(intBracketFlag<2) {
						switch(chrChemSyntaxString[intCharCount+1]) {
	                    //Europium
		                case 'u':
							intPast_Atom[intAttach_Cnt]=63;
							intQuantKey.Num_Etc++;
							break;
			            //Erbium
				        case 'r':
							intPast_Atom[intAttach_Cnt]=68;
							intQuantKey.Num_Etc++;
							break;
					    //Einsteinium
						case 's':
							intPast_Atom[intAttach_Cnt]=99;
							intQuantKey.Num_Etc++;
							break;
						}
					}
					if(intBracketFlag) intBracketFlag++;
					break;
               //Elements beginning with 'D'
               //Dysprosium
               case 'D':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=66;
						intQuantKey.Num_Etc++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Elements beginning with 'U'
               //Uranium
               case 'U':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=92;
						intQuantKey.Num_Etc++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Aryl carbon atom
               case 'c':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=6;
						intQuantKey.Num_AC++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Aryl oxygen atom
               case 'o':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=8;
						intQuantKey.Num_AO++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Aryl nitrogen atom
               case 'n':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=7;
						intQuantKey.Num_AN++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Aryl sulphur atom
               case 's':
                    if(intBracketFlag<2) {
						intPast_Atom[intAttach_Cnt]=16;
						intQuantKey.Num_AS++;
					}
					if(intBracketFlag)intBracketFlag++;
					break;
               //Aryl wild card atom
               case 'a':
                    break;
               //Double bond
               case '=':
                    intCurrentBond=2;
                    intQuantKey.Num_DB++;
                    break;
               //Triple bond
               case '#':
                    intCurrentBond=3;
                    intQuantKey.Num_TB++;
                    break;
               //Wild card atom (any atom)
               case '?':
                    break;
               default:
                    //Number of rings (Do not consider #'s used in brackets to denote # of hydrogens.)
                    if((chrChemSyntaxString[intCharCount]>48 && chrChemSyntaxString[intCharCount] < 58) && !intBracketFlag) {
                         intRing_Number=Get_SMILES_RN(intCharCount); 
						 if(intRingArray[intRing_Number]) {
                              intQuantKey.Num_Rings++;
                              intRingArray[intRing_Number]=0;
                         }
						 intRingArray[intRing_Number]=intRing_Number;
                    }
               }
                                   
          }
          
          //Default set the intCurrentBond flag to 1 (single bond) if it is set to other than a single bond and current character is an aliphatic atom (capital letter).
          //Note that aryl atoms are not considered because aryl atoms cannot be double bonded to aliphatic atoms. Also this is ok because capital letters in MF operators, etc. are skipped
          //over in previous loop block.
          if(chrChemSyntaxString[intCharCount] >64 && chrChemSyntaxString[intCharCount] < 91) intCurrentBond=1;

          //Reset Operator Flags for single atom occurrences (no group flag set).
          //Grouping braces {,} are associated w/ at least one of the current operators.
		  if(intGF_BF) {
			  while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))&& (chrChemSyntaxString[intCharCount] != '{' && chrChemSyntaxString[intCharCount] != '[')) {
				  //If despecification operator is associated with this combo bracket, then set back operator flags.
                    if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
                         intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
                         intOF_Que[intOF_Cntr]=0;
                         if(intOF_Cntr)intOF_Cntr--;
                    }
			  }
		  }
		  //No grouping braces {,} are associated w/ current operators.
		  else if(chrChemSyntaxString[intCharCount] != '{' && chrChemSyntaxString[intCharCount] != '[') {
			  intOF_BF=0;
			  intOF_Cntr=0;
			  intOF_Que[intOF_Cntr]=0;
		  }
		
          //Switch statment for decharacterization and unspecification characters
          switch(chrChemSyntaxString[intCharCount]) {
          
		  //Opening a branch
		  case '(':
			   intAttach_Cnt++;
			   intPast_Atom[intAttach_Cnt]=intPast_Atom[intAttach_Cnt-1];
			   break;
		  //Closing a branch
		  case ')':
			   intPast_Atom[intAttach_Cnt]=0;
			   intAttach_Cnt--;
			   break;
		  //Not operator (serves as a constraining despecifier.
          case '!':
			   switch(chrChemSyntaxString[intCharCount+1]) {
			   //Not operator is associated with a bond.
			   case '#':
			   case '=':
			   case '-':
			   case ':':
					intCharCount++;
					break;
  			   default:
					//Not operator is associated with an atom or subfragment.
					intOF_Cntr++;
					intOF_Que[intOF_Cntr]=2;
					intOF_BF |= (1<<intOF_Que[intOF_Cntr]);
			   }	
			   break;
          //Set decharacterization flag if character is a decharacterization character. Decharacterized characters are not counted for pre-screening.
          case '*':
               intOF_Cntr++;
               intOF_Que[intOF_Cntr]=1;
               intOF_BF |= (1<<intOF_Que[intOF_Cntr]);
               break;
          //Beginning of compound MF operator
		  case '<':
			   while(!(chrChemSyntaxString[intCharCount]=='>' && chrChemSyntaxString[intCharCount-1]=='>')) {
					intCharCount++;
			   }
			   break; 
		  //End of compound MF operator
          case '>':
//               intOF_Cntr++;
//               intOF_Que[intOF_Cntr]=3;
//               intOF_BF |= (1<<intOF_Que[intOF_Cntr]);
               break;
          //Set grouping brackets flag
          case '{':
               intGF_BF |= (1<<intOF_Que[intOF_Cntr]);
               break;
          //Clear grouping brackets flag
          case '}':
               if(intGF_BF & (1<<intOF_Que[intOF_Cntr])) intGF_BF ^=(1<<intOF_Que[intOF_Cntr]);
			   //Grouping braces {,} are associated w/ at least one of the current operators.
			   if(intGF_BF) {
					while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))) {
						//If despecification operator is associated with this combo bracket, then set back operator flags.
						if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
							intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
							intOF_Que[intOF_Cntr]=0;
							if(intOF_Cntr)intOF_Cntr--;
						}
					}
				} 
			    break;
          //Set atom bracket flag
          case '[':
               //Check to see if listing in brackets is a combo [X,x] or an atom/atom complex [SiH] by looking for a comma character.
               //Note: This is where you would insert code to analyze whether a bracketed atom possesses correct number of hydrogens.
               intCharTempCnt=intCharCount;
               intComboFlag=0;
               //Look for a comma.
               while(chrChemSyntaxString[intCharTempCnt] != ']') {
                    //If comma, then set combo flag.
                    if(chrChemSyntaxString[intCharTempCnt]==',') intComboFlag=1;
                    intCharTempCnt++;
               }
               //If a comma(combo) was found, then set the intCharCount variable to the end of the brackets.
               if(intComboFlag) {
                    intCharCount=intCharTempCnt;
                    
			        //Grouping braces {,} are associated w/ at least one of the current operators.
				    if(intGF_BF) {
						while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))) {
							//If despecification operator is associated with this combo bracket, then set back operator flags.
							if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
								intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
								intOF_Que[intOF_Cntr]=0;
								if(intOF_Cntr)intOF_Cntr--;
							}
						}
					}
					//No grouping braces {,} are associated w/ current operators.
					else {
						intOF_BF=0;
						intOF_Cntr=0;
						intOF_Que[intOF_Cntr]=0;
					}
		            intNumberAtoms++;
               }
               //Begins a non-combo atom bracket grouping.
               else {
                    //If despecifier operator is present, set grouping flag.
                    if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) intGF_BF |= (1<<intOF_Que[intOF_Cntr]);
                    //Set bracket flag     
                    intBracketFlag=1;
               }
               break;
          //Clear atom bracket and decharacterization flag if group (curly braces) flag is not set
          case ']':
               intBracketFlag=0;
               
			   if(intGF_BF & (1<<intOF_Que[intOF_Cntr])) intGF_BF ^=(1<<intOF_Que[intOF_Cntr]);
			   //Grouping braces {,} are associated w/ at least one of the current operators.
			   if(intGF_BF) {
					while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))) {
						//If despecification operator is associated with this combo bracket, then set back operator flags.
						if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
							intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
							intOF_Que[intOF_Cntr]=0;
							if(intOF_Cntr)intOF_Cntr--;
						}
					}
				}
				break;
          //Aryl carbon atom
          case 'c':
               if(intBracketFlag !=1) intNumberAtoms++;
               break;
          //Aryl oxygen atom
          case 'o':
               if(intBracketFlag !=1) intNumberAtoms++;
               break;
          //Aryl nitrogen atom
          case 'n':
               if(intBracketFlag !=1) intNumberAtoms++;
               break;
          //Aryl sulphur atom
          case 's':
               if(intBracketFlag !=1) intNumberAtoms++;
               break;
          //Aryl wild card atom
          case 'a':
               if(intBracketFlag !=1) intNumberAtoms++;
               break;
          //Wild card atom (any atom)
          case '?':
                intNumberAtoms++;
                break;
          default:
               //If character is a capital letters.
               if(chrChemSyntaxString[intCharCount] < 91 && chrChemSyntaxString[intCharCount] > 64) intNumberAtoms++;
          }
          //Increment the character counter
          intCharCount++;
     }
}

//This routing fills the atom occupation keys from the quantity keys. The routine is for two-tier screening if file option is selected.  This is for creating sort buckets for file searches.
void 
ChemSeqID::QK_to_OK() {
//   intOccupancyKey mapping: |10 | 9 |  8    |   7    | 6 |   5    | 4 |   3    |   2  |  1   |    0    |
//                      | C | c | s,o,n | rings? | O | DB (=) | N | TB (#) | Cl,I | F,Br | Etc,S,P |
     
     //Miscellaneous atoms, Sulfur, and Phosphorous
     if(intQuantKey.Num_Etc || intQuantKey.Num_S || intQuantKey.Num_P) intOccupancyKey |= (1<<0);
     //Fluorine and Bromine
     if(intQuantKey.Num_F || intQuantKey.Num_Br) intOccupancyKey |= (1<<1);
     //Chlorine and Iodine
     if(intQuantKey.Num_Cl || intQuantKey.Num_I) intOccupancyKey |= (1<<2);
     //Triple Bonds
     if(intQuantKey.Num_TB) intOccupancyKey |= (1<<3);
     //Aliphatic Nitrogen
     if(intQuantKey.Num_N) intOccupancyKey |= (1<<4);
     //Double Bonds
     if(intQuantKey.Num_DB) intOccupancyKey |= (1<<5);
     //Aliphatic Oxygen
     if(intQuantKey.Num_O) intOccupancyKey |= (1<<6);
     //Ring Structures
     if(intQuantKey.Num_Rings) intOccupancyKey |= (1<<7);
     //Aromatic Sulfur, Oxygen and Nitrogen
     if(intQuantKey.Num_AS || intQuantKey.Num_AO || intQuantKey.Num_AN) intOccupancyKey |= (1<<8);
     //Aromatic Carbon
     if(intQuantKey.Num_AC) intOccupancyKey |= (1<<9);
     //Aliphatic Carbon
     if(intQuantKey.Num_C) intOccupancyKey |= (1<<10);
}

//Fill_AID_SID function fills the bitfields, Atom_ID and Search_ID, for each Atom after initial parsing.
//Filling is accomplished via the ADA. The initial string is no longer necessary.
//This provides a descriptive bitfield for each atom used in the Subisomorphism routine.
void
ChemSeqID::Fill_AID_SID() {
     
	 AtomBond *ptrTemp=0;                    //Temporary pointer address
     unsigned _int32 intAtomCounter;         //Loop atom counter
     unsigned _int8 intNumAttachments=0;     //Number of atom attachments counter
     
     //Loop through Atom array
     for(intAtomCounter=0;intAtomCounter < intNumberAtoms; intAtomCounter++) {

          //If atom character is not "decharacterized"
          if(!(ptrMolecule->ptrAtom[intAtomCounter].Search_ID.Dechar_Atom)) {

               ptrTemp=ptrMolecule->ptrAtom[intAtomCounter].ptrNextBond;
          
               //Loop through all attachments at respective Atom element
               while(ptrTemp) {
               
                    intNumAttachments++;

                    //If this connection is not an unspecified atom or a not operator specified atom (essentially unspecified).
                    if(!ptrMolecule->ptrAtom[ptrTemp->intAttachedAtom].Search_ID.Unspec_Atom && !ptrMolecule->ptrAtom[ptrTemp->intAttachedAtom].Search_ID.Not_Atom){

                         //Set atom attachment type
                         switch(ptrMolecule->ptrAtom[ptrTemp->intAttachedAtom].Atom_ID.Atom_Type) {
                         //Hydrogen
                         case 1:
                              break;
                         //Helium
                         case 2:
                         //Lithium
                         case 3:
                         //Beryllium
                         case 4:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<4);
                              break;
                         //Boron
                         case 5:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<1);
                              break;
                         //Carbon
                         case 6:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<0);
                              break;
                         //Nitrogen
                         case 7:   
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<1);
                              break;
                         //Oxygen
                         case 8:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<2);
                              break;
                         //Fluorine
                         case 9:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<3);
                              break;
                         //Neon
                         case 10:
                         //Sodium
                         case 11:
                         //Magnesium
                         case 12:
                         //Aluminum
                         case 13:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<4);
                              break;
                         //Silicon
                         case 14:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<0);
                              break;
                         //Phosphorous
                         case 15:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<1);
                              break;
                         //Sulfur
                         case 16:     
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<2);
                              break;
                         //Chlorine
                         case 17:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<3);
                              break;
                         //Argon
                         case 18:
                         //Potassium
                         case 19:
                         //Calcium
                         case 20:
                         //Scandium
                         case 21:
                         //Titanium
                         case 22:
                         //Vanadium
                         case 23:
                         //Chromium
                         case 24:
                         //Manganese
                         case 25:
                         //Iron
                         case 26:
                         //Cobalt
                         case 27:
                         //Nickel
                         case 28:
                         //Copper
                         case 29:
                         //Zinc
                         case 30:
                         //Gallium
                         case 31:
                         //Germanium
                         case 32:
                         //Arsenic
                         case 33:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<4);
                              break;
                         //Selenium
                         case 34:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<2);
                              break;
                         //Bromine
                         case 35:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<3);
                              break;
                         //Krypton
                         case 36:
                         //Rubidium
                         case 37:
                         //Strontium
                         case 38:
                         //Yttrium
                         case 39:
                         //Zirconium
                         case 40:
                         //Niobium
                         case 41:
                         //Molybdenum
                         case 42:
                         //Technetium
                         case 43:
                         //Ruthenium
                         case 44:
                         //Rhodium
                         case 45:
                         //Palladium
                         case 46:
                         //Silver
                         case 47:
                         //Cadmium
                         case 48:
                         //Indium
                         case 49:
                         //Tin
                         case 50:
                         //Antimony
                         case 51:
                         //Tellurium
                         case 52:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<4);
                              break;
                         //Iodine
                         case 53:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<3);
                              break;
                         //Xenon
                         case 54:
                         //Cesium
                         case 55:
                         //Barium
                         case 56:
                         //Lanthanum
                         case 57:
                         //Cerium
                         case 58:
                         //Praseodymium
                         case 59:
                         //Neodymium
                         case 60:
                         //Promethium
                         case 61:
                         //Samarium
                         case 62:
                         //Europium
                         case 63:
                         //Gadolinium
                         case 64:
                         //Terbium
                         case 65:
                         //Dysprosium
                         case 66:
                         //Holmium
                         case 67:
                         //Erbium
                         case 68:
                         //Thulium
                         case 69:
                         //Ytterbium
                         case 70:
                         //Lutetium
                         case 71:
                         //Hafnium
                         case 72:
                         //Tantalum
                         case 73:
                         //Wolfram
                         case 74:
                         //Rhenium
                         case 75:
                         //Osmium
                         case 76:
                         //Iridium
                         case 77:
                         //Platinum
                         case 78:
                         //Gold
                         case 79:
                         //Mercury
                         case 80:
                         //Thallium
                         case 81:
                         //Lead
                         case 82:
                         //Bismuth
                         case 83:
                         //Polonium
                         case 84:
                         //Astatine
                         case 85:
                         //Radon
                         case 86:
                         //Francium
                         case 87:
                         //Radium
                         case 88:
                         //Actinum
                         case 89:
                         //Thorium
                         case 90:
                         //Protactinium
                         case 91:
                         //Uranium
                         case 92:
                         //Neptunium
                         case 93:
                         //Plutonium
                         case 94:
                         //Americum
                         case 95:
                         //Curium
                         case 96:
                         //Berkelium
                         case 97:
                         //Californium
                         case 98:
                         //Einsteinium
                         case 99:
                         //Fermium
                         case 100:
                         //Mendelevium
                         case 101:
                         //Nobelium
                         case 102:
                         //Lawrencium
                         case 103:
                              ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Atom_Cnt |= (1<<4);
                              break;
                         }
                    }
                    //Else if the connection is an unspecified atom, then set unspecifed atom connection flag
                    else {
                         //Set unspecified atom connection flag.
                         ptrMolecule->ptrAtom[intAtomCounter].Search_ID.Unspec_Neighbor=1;
                    }
                    
                    //Determine bonding environment.
                    switch(ptrTemp->Bond_ID.Bond_Type) {
                    //Single bond attachment 
                    case 1:
                         ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Bond_Types |=(1<<0);
                         break;
                    //Double bond attachment
                    case 2:
                         ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Bond_Types |=(1<<1);
                         break;
                    //Triple bond attachment
                    case 3:
                         ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Bond_Types |=(1<<2);
                         break;
                    //Aromatic (delocalized) bond
                    case 4:
                         ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Bond_Types |=(1<<3);
                         break;
                    //Wild (any) bond
                    case 15:
                         ptrMolecule->ptrAtom[intAtomCounter].Search_ID.Unspec_Bond =1;
                         break;
                    }
                    ptrTemp=ptrTemp->ptrNextBond;
               }
               ptrMolecule->ptrAtom[intAtomCounter].Atom_ID.Num_NH_Cnt=intNumAttachments;
               intNumAttachments=0;
          }
     }
}

unsigned _int16 
ChemSeqID::Get_SMILES_RN(unsigned _int32 &intCharCount) {

unsigned _int16 intRingDigit=0;         //Returned ring digit
unsigned _int32 intM_Digit=0;           //Flag/storage for presence/location of multiple digit demarcation (%)
unsigned _int32 intTempCntr;           //Temporary counter
unsigned _int8 i=0;                     //Exponent counter

intTempCntr=intCharCount;

while(chrChemSyntaxString[intTempCntr] > 48 && chrChemSyntaxString[intTempCntr] < 58) {
     if(chrChemSyntaxString[intTempCntr]== '%') {
          intM_Digit=intTempCntr;
          break;
     }
     intTempCntr++;
}
i=0;
if(intM_Digit) {
     while(intCharCount != intTempCntr) {
          intRingDigit += (chrChemSyntaxString[intCharCount]-48)*pow(10,i);
          intCharCount++;
     }
}
else intRingDigit=(chrChemSyntaxString[intCharCount]-48);

return intRingDigit;

}

//This routine parses a SMILES string into an adjacency listing of the molecular structure. Then it calls the structure filling routine to
//fill in the structure ID bitfields with information regarding the atomic environment from the adjacency listing.
void 
ChemSeqID::Parse_SMILES(){

     //Initial variable, pointer and array declarations.
	 unsigned _int8 i;
     unsigned _int32 intCharCount=0;         //Current location in SMILES string array
     unsigned _int32 intCharTempCnt=0;		 //Temporary SMILES string counter
	 unsigned _int32 intAtom_Node=0;         //New location in Atom array
     unsigned _int8 intAttachCount=0;        //Number of branches currently open
     unsigned _int8 intBondValue=1;          //Current value of open bond  | 15~wild | 4~aromatic | 3~triple | 2~double | 1~single |
	 unsigned _int8 intNotBond=0;			 //Flag denoting whether current bond is logically "not"ted.
     unsigned _int8 intBracketFlag=0;        //Flag denoting brackets []. Also serves as a counter for # of atoms in the brackets.
     unsigned _int8 intComboFlag=0;			 //Flag determining whether current atom character is in a combo (i.e., [C.c]) or a single atom (0~no combo,1~combo)
	 unsigned _int16 intRingNumber=0;		 //Integer location numbe assigned to SMILES ring closure
     unsigned _int8 intOF_Cntr=0;            //Molecular operation flag counter
	 unsigned _int8 intFeatureFlag=0;        //  ...| aryl ring | alicyclic ring | ring bond | non-ring bond |   
     unsigned _int8 intOF_BF=0;              //Flag bit field for molecular operators  ...| 3~MF operator | 2~Not atom | 1~Dechar atom |0~blank | 
     unsigned _int8 intGF_BF=0;              //Flag bit filed for grouping flags
     unsigned _int8 intOF_Que[5]={0};        //Que array storing type of despecification operator at appropriate level of occurrence
     SearchID Temp_SID;                      //Temporary search ID used with compound Molecular Feature operators
     
     Atom *ptrStructure;                     //Temporary atom pointer
     AtomBond *ptrRingAttach[60]={0};		 //Array of bond connection locations in Atom array of ring closures
	 _int32 intRingLoc[60];                 //Array of integer locations for ring closures.  Currently supports only 10 open rings at a time.
     unsigned _int8 intRingBond[60]={0};     //Array of ring closure bond type (value)   
     AtomBond *ptrBranchAttach[100]={0};     //Array of bond connection locations in Atom array of nested branches.  Currently 100 max.
     _int32 intBranchLoc[100]={0};           //Array of integer bond connection locations within the adjacency listing spine

	 //Initialize intRingLoc to -1
	 for (i=0;i<60;i++) {
          intRingLoc[i]=-1;
	 }
	 
     //Create a molecule.
	 if(!ptrMolecule) ptrMolecule=new Molecule;

	 //Parse alpha (master) structure from SMILES nomenclature to adjacency listing
     if(!ptrMolecule->ptrAtom) ptrMolecule->ptrAtom=new Atom[intNumberAtoms];
     ptrStructure=ptrMolecule->ptrAtom;

     //Loop to traverse characters in SMILES string while simultaneously constructing a dynamically allocated linked list 
     //representing the molecular graph.  (Linked list resembles an adjacency listing).
     while(chrChemSyntaxString[intCharCount] != '\0') {
          
          switch(chrChemSyntaxString[intCharCount]) {
          //Elements (atoms) beginning with 'H'
          case 'H':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Helium
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =2;
                    break;
               //Hafnium
               case 'f':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =72;
                    break;
               //Mercury
               case 'g':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =80;
                    break;
               //Holmium
               case 'o':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =67;
                    break;
               //Hydrogen
               default:
				    break;
               }
			   if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'L'
          case 'L':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Lithium
               case 'i':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =3;
                    break;
               //Lanthanum
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =57;
                    break;
               //Lutetium
               case 'u':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =71;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'B'
          case 'B':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Beryllium
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =4;
                    break;
               //Bromine
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =35;
                    break;
               //Barium
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =56;
                    break;
               //Bismuth
               case 'i':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =83;
                    break;
               //Berkelium
               case 'k':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =97;
                    break;
               //Boron
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =5;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'C'
          case 'C':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Chlorine
               case 'l':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =17;
                    break;
               //Calcium
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =20;
                    break;
               //Chromium
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =24;
                    break;
               //Cobalt
               case 'o':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =27;
                    break;
               //Copper
               case 'u':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =29;
                    break;
               //Cadmium
               case 'd':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =48;
                    break;
               //Cesium
               case 's':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =55;
                    break;
               //Cerium
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =58;
                    break;
               //Curium
               case 'm':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =96;
                    break;
               //Carbon
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =6;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'N'
          case 'N':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Neon
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =10;
                    break;
               //Sodium
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =11;
                    break;
               //Nickel
               case 'i':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =28;
                    break;
               //Niobium
               case 'b':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =41;
                    break;
               //Neodymium
               case 'd':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =60;
                    break;
               //Neptunium
               case 'p':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =93;
                    break;
               //Nobellium
               case 'o':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =102;
                    break;
               //Nitrogen
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =7;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Element beginning with 'O'
          case 'O':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Osmium
               case 's':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =76;
                    break;
               //Oxygen
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =8;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'F'
          case 'F':    
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Iron
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =26;
                    break;
               //Francium
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =87;
                    break;
               //Fermium
               case 'm':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =100;
                    break;
               //Fluorine
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =9;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'M'
          case 'M':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Magnesium
               case 'g':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =12;
                    break;
               //Manganese
               case 'n':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =25;
                    break;
               //Molybdenum
               case 'o':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =42;
                    break;
               //Mendelevium
               case 'd':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =101;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'A'
          case 'A':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Aluminum
               case 'l':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =13;
                    break;
               //Argon
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =18;
                    break;
               //Arsenic
               case 's':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =33;
                    break;
               //Silver
               case 'g':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =47;
                    break;
               //Gold
               case 'u':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =79;
                    break;
               //Astatine
               case 't':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =85;
                    break;
               //Actinum
               case 'c':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =89;
                    break;
               //Americum
               case 'm':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =95;
                    break;
               //Aliphatic wild card atom
               default:
                    //Check to make sure that unspecified aliphatic atom is not preceded by an MF operator.
                    //If it is, MF operator overrides unspecified aliphatic atom properties.
                    if(!ptrStructure[intAtom_Node].Search_ID.In_Ring) {
                       ptrStructure[intAtom_Node].Search_ID.Unspec_Atom=1;
					   ptrStructure[intAtom_Node].Search_ID.In_Ring=1;
                       ptrStructure[intAtom_Node].Search_ID.Not_Ring=1;
                       ptrStructure[intAtom_Node].Search_ID.Ring_Type |= (1<<0);
                    }
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'S'
          case 'S':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Silicon
               case 'i':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =14;
                    break;
               //Scandium
               case 'c':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =21;
                    break;
               //Selenium
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =34;
                    break;
               //Strontium
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =38;
                    break;
               //Tin
               case 'n':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =50;
                    break;
               //Antimony
               case 'b':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =51;
                    break;
               //Samarium
               case 'm':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =62;
                    break;
               //Sulfur
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =16;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'P'
          case 'P':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Paladium
               case 'd':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =46;
                    break;
               //Platinum
               case 't':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =78;
                    break;
               //Lead
               case 'b':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =82;
                    break;
               //Polonium
               case 'o':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =84;
                    break;
               //Praseodymium
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =59;
                    break;
               //Promethium
               case 'm':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =61;
                    break;
               //Protactinum
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =91;
                    break;
               //Plutonium
               case 'u':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =94;
                    break;
               //Phosphorous
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =15;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'K'
          case 'K':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Krypton
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =36;
                    break;
               //Potassium
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =19;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'T'
          case 'T':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Titanium
               case 'i':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =22;
                    break;
               //Technetium
               case 'c':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =43;
                    break;
               //Tellurium
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =52;
                    break;
               //Tantalum
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =73;
                    break;
               //Thallium
               case 'l':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =81;
                    break;
               //Terbium
               case 'b':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =65;
                    break;
               //Thorium
               case 'h':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =90;
                    break;
               //Thulium
               case 'm':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =69;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'V'
          //Vanadium
          case 'V':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type =23;
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'Z'
          case 'Z':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Zinc
               case 'n':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =30;
                    break;
               //Zirconium
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =40;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'G' 
          case 'G':     
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Gallium
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =31;
                    break;
               //Germanium
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =32;
                    break;
               //Gadolinium
               case 'd':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =64;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'R'
          case 'R':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Rubidium
               case 'b':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =37;
                    break;
               //Ruthenium
               case 'u':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =44;
                    break;
               //Rhodium
               case 'h':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =45;
                    break;
               //Rhenium
               case 'e':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =75;
                    break;
               //Radon
               case 'n':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =86;
                    break;
               //Radium
               case 'a':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =88;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'Y'
          case 'Y':
               switch(chrChemSyntaxString[intCharCount+1]){
               //Ytterbium
               case 'b':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =70;
                    break;
               //Yttrium
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =39;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'I'
          case 'I':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Indium
               case 'n':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =49;
                    break;
               //Iridium
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =77;
                    break;
               //Iodine
               default:
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =53;
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'X'
          //Xenon
          case 'X':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type =54;
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'W'
          //Wolfram
          case 'W':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type =74; 
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'E'
          case 'E':
               switch(chrChemSyntaxString[intCharCount+1]) {
               //Europium
               case 'u':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =63; 
                    break;
               //Erbium
               case 'r':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =68; 
                    break;
               //Einsteinium
               case 's':
                    ptrStructure[intAtom_Node].Atom_ID.Atom_Type =99; 
                    break;
               }
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'D'
          //Dysprosium
          case 'D':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type =66;
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Elements beginning with 'U'
          //Uranium
          case 'U':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type =92;
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Aryl carbon atom
          case 'c':
			   ptrStructure[intAtom_Node].Atom_ID.Atom_Type=6;
               ptrStructure[intAtom_Node].Search_ID.In_Ring=1;
               ptrStructure[intAtom_Node].Search_ID.Ring_Type |= (1<<1); 
			   
			   if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node){
							if(ptrStructure[intBranchLoc[intAttachCount]].Search_ID.Ring_Type & (1<<1)) intBondValue=4;
							if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						}
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Aryl oxygen atom
          case 'o':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type=8;
               ptrStructure[intAtom_Node].Search_ID.In_Ring=1;
               ptrStructure[intAtom_Node].Search_ID.Ring_Type |= (1<<1);
               
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node){
							if(ptrStructure[intBranchLoc[intAttachCount]].Search_ID.Ring_Type & (1<<1)) intBondValue=4;
							if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						}
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Aryl nitrogen atom
          case 'n':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type=7;
               ptrStructure[intAtom_Node].Search_ID.In_Ring=1;
               ptrStructure[intAtom_Node].Search_ID.Ring_Type |= (1<<1);
               
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node){
							if(ptrStructure[intBranchLoc[intAttachCount]].Search_ID.Ring_Type & (1<<1)) intBondValue=4;
							if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						}
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Aryl sulphur atom
          case 's':
               ptrStructure[intAtom_Node].Atom_ID.Atom_Type=16;
               ptrStructure[intAtom_Node].Search_ID.In_Ring=1;
               ptrStructure[intAtom_Node].Search_ID.Ring_Type |= (1<<1);
               
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node){
							if(ptrStructure[intBranchLoc[intAttachCount]].Search_ID.Ring_Type & (1<<1)) intBondValue=4;
							if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						}
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Aryl wild card atom
          case 'a':
               ptrStructure[intAtom_Node].Search_ID.Unspec_Atom=1;
               ptrStructure[intAtom_Node].Search_ID.In_Ring=1;
               ptrStructure[intAtom_Node].Search_ID.Ring_Type |= (1<<1);
               
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node){
							if(ptrStructure[intBranchLoc[intAttachCount]].Search_ID.Ring_Type & (1<<1)) intBondValue=4;
							if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						}
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Unspecified (wild card atom/any atom) atom
          case '?':
               ptrStructure[intAtom_Node].Search_ID.Unspec_Atom=1;
               ptrStructure[intAtom_Node].Search_ID.Not_Ring=1;
               ptrStructure[intAtom_Node].Search_ID.In_Ring=1;
               ptrStructure[intAtom_Node].Search_ID.Ring_Type |= (2<<0);
               if(intAtom_Node < intNumberAtoms) {
					//If atom character is first atom in an atom bracket [...]
					if(intBracketFlag<2) { 
						//If search operator flags are present, then set appropriate flags in Search_ID bit field.
						Set_SOF_BF(intOF_Cntr,intOF_BF,intGF_BF,intOF_Que,ptrStructure[intAtom_Node],intCharCount,Temp_SID);
						if(intAtom_Node){
							if(intAtom_Node) AttachNewAtom(ptrStructure,intAtom_Node,ptrBranchAttach,intBranchLoc,intAttachCount,intBondValue,intNotBond);
						}
						intAtom_Node++;
					}
					if(intBracketFlag) intBracketFlag++;
			   }
               break;
          //Single bond
          case '-':
               intBondValue=1;
               break;
          //Double bond
          case '=':
               intBondValue=2; 
               break;
          //Triple bond
          case '#':
               intBondValue=3; 
               break;
          //Aromatic bond
          case ':':
               intBondValue=4;
               break;
          //Unspecified (wild card) bond
          case '~':
               intBondValue=15;
               ptrStructure[intAtom_Node].Search_ID.Unspec_Bond=1;
               break;
          //Open branch character
          case '(':
               intAttachCount++;
               ptrBranchAttach[intAttachCount]=ptrBranchAttach[intAttachCount-1];
               intBranchLoc[intAttachCount]=intBranchLoc[intAttachCount-1];
               break;
          //Close branch character
          case ')':
               ptrBranchAttach[intAttachCount]=0;
               intBranchLoc[intAttachCount]=0;
               intAttachCount--;
               if(ptrBranchAttach[intAttachCount]->ptrNextBond)ptrBranchAttach[intAttachCount]=ptrBranchAttach[intAttachCount]->ptrNextBond;
			   break;
          //Not operator (serves as a constraining despecifier.
          case '!':
			   switch(chrChemSyntaxString[intCharCount+1]) {
			   //Not operator is associated with a bond.
			   case '#':
			   case '=':
			   case '-':
			   case ':':
					intNotBond=1;
					break;
  			   default:
					//Not operator is associated with an atom or subfragment.
					intOF_Cntr++;
					intOF_Que[intOF_Cntr]=2;
					intOF_BF |= (1<<intOF_Que[intOF_Cntr]);
			   }
			   break;
          //Set decharacterization flag if character is a decharacterization character. Decharacterized characters are not counted for pre-screening.
          case '*':
               intOF_Cntr++;
               intOF_Que[intOF_Cntr]=1;
               intOF_BF |= (1<<intOF_Que[intOF_Cntr]);
               break;
          //Beginning of compound MF operator
          case '<':
			   //Increment operator que and flags for MF operator.
               intOF_Cntr++;
               intOF_Que[intOF_Cntr]=3;
               intOF_BF |= (1<<intOF_Que[intOF_Cntr]);
			  
			   //Check for feature commands located inside ' <<___>> '
			   intFeatureFlag=0;
							  
			   //Check for feature commands located inside ' <<___>> '
               if(chrChemSyntaxString[intCharCount+1] =='<') { 
                    
					//Initialize bitfields
                    intFeatureFlag=0;
                    Temp_SID.Reset_ID();
                    intCharCount +=2;
               
					while(!(chrChemSyntaxString[intCharCount] == '>' && chrChemSyntaxString[intCharCount-1] == '>')){
                         switch(chrChemSyntaxString[intCharCount]) {
                         //Not operator for use with Molecular Feature declarations
                         case '!':
                              if(chrChemSyntaxString[intCharCount+1]=='R' || chrChemSyntaxString[intCharCount+1]=='r') Temp_SID.Not_Ring=1;
                              break;
                         //Flag for # of alicyclic rings
                         case 'R':
                              Temp_SID.Ring_Type |= (1<<0);   
                              if(!Temp_SID.Not_Ring) Temp_SID.In_Ring=1;
                              intFeatureFlag |= (1<<2);
                              break;
                         //Flag for # of aryl rings
                         case 'r':
                              Temp_SID.Ring_Type |= (1<<1);   
                              if(!Temp_SID.Not_Ring) Temp_SID.In_Ring=1;
                              intFeatureFlag |= (1<<3);
                              break;
                         //Less than symbol (for ring search)
                         case '<':
                              if(intFeatureFlag & (3<<2)) Temp_SID.Less_Than=1;
                              break;
                         //Greater than symbol
                         case '>':
                              if(intFeatureFlag & (3<<2)) Temp_SID.Greater_Than=1;
                              break;
                         //Numbers used to specify bonds and ring sizes
                         default:
                              if(chrChemSyntaxString[intCharCount] >48 && chrChemSyntaxString[intCharCount] <58) {
                                   switch(intFeatureFlag) {
                                   //Aryl and Alicyclic rings
                                   case 12:
                                   //Aryl ring
                                   case 8:
                                        Temp_SID.Ring_Size = (chrChemSyntaxString[intCharCount]-48);
                                        break;
                                   //Alicyclic ring
                                   case 4:
                                        Temp_SID.Ring_Size = (chrChemSyntaxString[intCharCount]-48);
                                        break;
                                   }
                              }
							  break;
                         }               
                         intCharCount++;
                    }
                    intCharCount++;
               }
               break;
          //Clear molecular feature constraint flag.
		  case '>':
			   if(intGF_BF & (1<<intOF_Que[intOF_Cntr])) intGF_BF ^=(1<<intOF_Que[intOF_Cntr]);
			   //Grouping braces {,} are associated w/ at least one of the current operators.
			   if(intGF_BF) {
					while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))) {
						//If despecification operator is associated with this combo bracket, then set back operator flags.
						if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
							intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
							intOF_Que[intOF_Cntr]=0;
							if(intOF_Cntr)intOF_Cntr--;
						}
					}
			   }
			   break; 
		  //Set grouping brackets flag
          case '{':
               intGF_BF |= (1<<intOF_Que[intOF_Cntr]);
               break;
          //Clear grouping brackets flag
          case '}':
               if(intGF_BF & (1<<intOF_Que[intOF_Cntr])) intGF_BF ^=(1<<intOF_Que[intOF_Cntr]);
			   //Grouping braces {,} are associated w/ at least one of the current operators.
			   if(intGF_BF) {
					while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))) {
						//If despecification operator is associated with this combo bracket, then set back operator flags.
						if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
							intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
							intOF_Que[intOF_Cntr]=0;
							if(intOF_Cntr)intOF_Cntr--;
						}
					}
				}
			    break;
          //Set atom bracket flag
          case '[':
               //Check to see if listing in brackets is a combo [X,x] or an atom/atom complex [SiH] by looking for a comma character.
               //Note: This is where you would insert code to analyze whether an atom possesses correct number of hydrogens.
               intCharTempCnt=intCharCount+1;
                              
			   //If despecifier operator is present, set grouping flag.
               if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) intGF_BF |= (1<<intOF_Que[intOF_Cntr]);
			   intBracketFlag=1;	//Set bracket flag     
               break;
          //Clear atom bracket and decharacterization flag if group (curly braces) flag is not set
          case ']':
               intBracketFlag=0;
               
			   //Clear gouping bit for current que level if set.
			   if(intGF_BF & (1<<intOF_Que[intOF_Cntr])) intGF_BF ^=(1<<intOF_Que[intOF_Cntr]);
			   //Grouping braces {,} are associated w/ at least one of the current operators.
			   if(intGF_BF) {
					while(!(intGF_BF & (1<<intOF_Que[intOF_Cntr]))) {
						//If despecification operator is associated with this combo bracket, then set back operator flags.
						if(intOF_BF & (1<<intOF_Que[intOF_Cntr])) { 
							intOF_BF ^= (1<<intOF_Que[intOF_Cntr]);
							intOF_Que[intOF_Cntr]=0;
							if(intOF_Cntr)intOF_Cntr--;
						}
					}
				}	
               break;
          default: 
               //Account for ring opening or closure. Remember must not be present inside an atom character bracket [,].
               if((chrChemSyntaxString[intCharCount] > 48 && chrChemSyntaxString[intCharCount] < 58) && !intBracketFlag) {
                    //Get ring number from SMILES string
                    intRingNumber=Get_SMILES_RN(intCharCount);

                    if(intRingLoc[intRingNumber]>=0) {
                         
						 ptrRingAttach[intRingNumber]=ptrStructure[intRingLoc[intRingNumber]].ptrNextBond;
						 AttachNewAtom(ptrStructure,intAtom_Node-1,ptrRingAttach,intRingLoc,intRingNumber,intRingBond[intRingNumber],intNotBond);

                         //Reset RingLoc[~] location to default for future reuse     
                         ptrRingAttach[intRingNumber]=0;
						 intRingLoc[intRingNumber]=-1;
                         intRingBond[intRingNumber]=0;
						 intBondValue=1;
                    }
					//If ring opening then save location
                    else {
                         intRingLoc[intRingNumber]=intBranchLoc[intAttachCount];
                         
						 if(ptrStructure[intBranchLoc[intAttachCount]].Search_ID.Ring_Type & (1<<1)) intRingBond[intRingNumber]=4;
						 else intRingBond[intRingNumber]=intBondValue;
						 
                         intBondValue=1;
                    }
               }
               break;
          }

          //Increment character count
          intCharCount++;
     }

	 Fill_AID_SID();
//*********************************************************
//**** Printout adjacency listing for debugging purposes
//unsigned _int32 counter;
//AtomBond *ptrTemp=0;
//for(counter=0;counter<intAtom_Node;counter++) {
//   cout<<" Atom: "<<ptrStructure[counter].Atom_ID.Atom_Type<<" Dechar: "<<ptrStructure[counter].Search_ID.Dechar_Atom<<" Unspec: "<<ptrStructure[counter].Search_ID.Unspec_Atom<<" Not: "<<ptrStructure[counter].Search_ID.Not_Atom<<" => ";
//   ptrTemp=ptrStructure[counter].ptrNextBond;
//   while(ptrTemp != 0) {
//        cout<<ptrTemp->Bond_ID.Bond_Type<<" "<<ptrTemp->intAttachedAtom<<"=>";
//        ptrTemp=ptrTemp->ptrNextBond;
//   }
//   cout<<endl;
//}
//cout<<endl;
//********************************************************* 

}
