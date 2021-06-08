///////////////////////////////////////////////////////////////
//	This is the function code accompanying the 
//	header file SDA_2.h for MOSDAP(c) v2.0
//
//	Copyright (c) 1998. John W. Raymond.  All rights reserved.
//
///////////////////////////////////////////////////////////////

#include "SDA_2.H"

QueNode::QueNode() {
	intAtom_Index=0;
	intSource=0;
	ptrNextQue=0;
}

QueNode::~QueNode() {
	if(ptrNextQue) delete ptrNextQue;
}

//Adds and fills a qeue node element to the qeue.
void
QueNode::Add_Que(unsigned _int32 intSource_Node,unsigned _int32 intIndex_Node) {

	ptrNextQue=new QueNode;
	ptrNextQue->intSource=intSource_Node;
	ptrNextQue->intAtom_Index=intIndex_Node;
}

HashHeader::HashHeader () {
     intChainLength=0;
     ptrSepChain=0;
}

HashHeader::~HashHeader() {
     if(ptrSepChain) delete[] ptrSepChain;
}

HashStructure::HashStructure() {
	intHash_PN=19;
	intHash_Flag=0;
	ptrHash_Bucket=new HashHeader[intHash_PN];
}

HashStructure::~HashStructure() {
	if(ptrHash_Bucket) delete[] ptrHash_Bucket;
}

MatchElement::MatchElement() {
     intMatch=0;
     ptrNextMatch=0;
}

MatchElement::~MatchElement() {
     if(ptrNextMatch) delete ptrNextMatch;
}

//Adds a match node to a given match node in a list.
void
MatchElement::Add_MN(unsigned _int32 intQuery_Loc){

	ptrNextMatch=new MatchElement;
	ptrNextMatch->intMatch=intQuery_Loc;
	
}

//Deletes a match node from a match element list.
unsigned _int8
MatchElement::Delete_MN(MatchElement *&ptrOld_Match) {
	MatchElement *ptrTempMatch;
	
	//If match element has an attachment.
	if(ptrNextMatch) {
		intMatch=ptrNextMatch->intMatch;
		ptrTempMatch=ptrNextMatch;

		ptrNextMatch=ptrNextMatch->ptrNextMatch;
		ptrTempMatch->ptrNextMatch=0;
		delete ptrTempMatch;
	}
	//Else if match element is the last node in the list.
	else {
		if(this != ptrOld_Match) {
			delete this;
			ptrOld_Match->ptrNextMatch=0;
		}
		else return 0;
	}
	return 1;
}

//This routine clears a hash table for another hashed graph.
void
HashStructure::Reset_HT() {

     _int8 intHash_Cntr;

     for(intHash_Cntr=0;intHash_Cntr<intHash_PN;intHash_Cntr++) {
          ptrHash_Bucket[intHash_Cntr].intChainLength=0;
          delete ptrHash_Bucket[intHash_Cntr].ptrSepChain;
          ptrHash_Bucket[intHash_Cntr].ptrSepChain=0;
     }
	 intHash_Flag=0;
}

//This routine hashes a graph represented as an adjancency multi-list into a hash table
void
HashStructure::Hash_Graph(Atom *ptrStructure,unsigned _int32 intNum_Atoms) {

     unsigned _int32 i;													//Loop counter
     unsigned _int32 *ptrSep_Chain_Loc=new unsigned _int32[intHash_PN];	//Temporary location in separate chain used in the filling process
     unsigned _int32 *ptrHashHeadLoc=new unsigned _int32[intNum_Atoms]; //Stores location of Atom in hash header to reduce the # of hashing operations to be performed

	 intHash_Flag=1;

	 //Initialize separate chain array.
	 for(i=0;i<intHash_PN;i++) {
		 ptrSep_Chain_Loc[i]=0;
	 }

     //Fill intChainLengths[~] vector for use in filling hash table
     for(i=0;i<intNum_Atoms;i++) {
		  if(ptrStructure[i].Search_ID.Dechar_Atom) ptrHashHeadLoc[i]=(ptrStructure[i].Atom_ID.Atom_Type % intHash_PN);
		  else ptrHashHeadLoc[i]=((ptrStructure[i].Atom_ID.Atom_Type | (ptrStructure[i].Atom_ID.Num_NH_Cnt<<7)) % intHash_PN);
		
          ptrHash_Bucket[ptrHashHeadLoc[i]].intChainLength++;
     }

     //Allocate the length of each separate chain to its respective hash header pointer
     for(i=0;i<intHash_PN;i++) {
          //Create separate chain links if the specified chain length is greater than zero
          if(ptrHash_Bucket[i].intChainLength) ptrHash_Bucket[i].ptrSepChain=new unsigned _int32[ptrHash_Bucket[i].intChainLength];
     }

     //Fill the hash table
     for(i=0;i<intNum_Atoms;i++) {
          //(If chain length is greater than zero)
          if(ptrHash_Bucket[ptrHashHeadLoc[i]].intChainLength) {
               *(ptrHash_Bucket[ptrHashHeadLoc[i]].ptrSepChain + ptrSep_Chain_Loc[ptrHashHeadLoc[i]])=i;
               ptrSep_Chain_Loc[ptrHashHeadLoc[i]]++;
          }
     }

     delete[] ptrHashHeadLoc;
	 delete[] ptrSep_Chain_Loc;
}

//Places a match element node onto the storage "rack".
unsigned _int8
MatchStructure::Rack_MN(unsigned _int32 intRack_Row,unsigned _int32 intRack_Depth,MatchElement *&ptrOld_Match,MatchElement *&ptrTempMatch) {

	//Decrement row and depth counters from the refinement procedure because "rack" storage structure has one less row and depth than number of atoms
	intRack_Row--;
	intRack_Depth--;

	//Place on rack via the last address storage structure.
	if(ptrRefine_Top[intRack_Row][intRack_Depth]){

		ptrRefine_Top[intRack_Row][intRack_Depth]->ptrNextMatch=ptrTempMatch;
		ptrRefine_Top[intRack_Row][intRack_Depth]=ptrRefine_Top[intRack_Row][intRack_Depth]->ptrNextMatch;
	}
	//First match element node in this location.
	else {
		ptrRefine_Rack[intRack_Row][intRack_Depth]=ptrTempMatch;
		ptrRefine_Top[intRack_Row][intRack_Depth]=ptrTempMatch;
	}

	//Remove match element node from match structure spine.
	if(ptrOld_Match != ptrTempMatch) {
		ptrOld_Match->ptrNextMatch=ptrTempMatch->ptrNextMatch;
		ptrTempMatch->ptrNextMatch=0;
	}
	else {
		if(ptrTempMatch->ptrNextMatch) {
			ptrOld_Match=ptrTempMatch->ptrNextMatch;
			ptrMS_Spine[intRack_Row+1]=ptrOld_Match;
			ptrTempMatch->ptrNextMatch=0;
		}
		else return 0;
	}
	return 1;
}

//Removes a vector of "stored" match element nodes from the "storage" structure and places back onto the spine array.
void
MatchStructure::UnRack_MN(unsigned _int32 intRack_Depth,unsigned _int32 intNum_Rack_Atoms) {

unsigned _int32 j;
MatchElement *ptrTempMatch;

for(j=intRack_Depth;j<intNum_Rack_Atoms-1;j++) {
	//Find end of list chain.
	ptrTempMatch=ptrMS_Spine[j+1];
	while(ptrTempMatch->ptrNextMatch) {
		ptrTempMatch=ptrTempMatch->ptrNextMatch;
	}

	ptrTempMatch->ptrNextMatch=ptrRefine_Rack[j][intRack_Depth];
	ptrRefine_Rack[j][intRack_Depth]=0;
	ptrRefine_Top[j][intRack_Depth]=0;
}

}

//Removes all vectors of "stored match element nodes from the "storage" structure and places back onto the spine array.
void
MatchStructure::UnRack_All(unsigned _int32 intNum_Rack_Atoms) {

unsigned _int32 i;

for(i=0;i<intNum_Rack_Atoms-1;i++) {

	UnRack_MN(i,intNum_Rack_Atoms);
}
}

//This routine truncates the match listing, removing all but depth zero and deleting detected atoms from depth zero, so that the same atoms are not used in additional substructure searching.
unsigned _int8
MatchStructure::Truncate_Spine(ChemSeqID *ptrAlpha,ChemSeqID *ptrBeta) {

unsigned _int32 intZeroDepth=0;			//Depth of initial match listing to be used in deletions
unsigned _int32 intBetaCntr;			//Counter variable
MatchElement *ptrOldTempMatch=0;		//Pointer preceeding delete location in MatchListing structure
MatchElement *ptrTempMatch=0;			//Pointer used in deleting location in MatchListing structure

for(intBetaCntr=0;intBetaCntr<ptrBeta->intNumberAtoms;intBetaCntr++) {

     //Make sure detected atom is not a decharacterized atom
     if(!ptrBeta->ptrMolecule->ptrAtom[intBetaCntr].Search_ID.Dechar_Atom) {
          
          //Operation block used to delete previously detected atoms in ptrColLocaion array from match structure so that additional occurences of subfragment may be found.
          //If current match element possesses a subsequent attachment, use member deletion function without passing in previous match element address (not needed in this case).
		  if(ptrCol_Loc[intBetaCntr]->ptrNextMatch) ptrCol_Loc[intBetaCntr]->Delete_MN(ptrOldTempMatch);
		  //Else determine previous match element location because current match element does not possess a subsequent attachment (previous address is needed in this case).
		  else {
			  ptrOldTempMatch=ptrMS_Spine[intBetaCntr];
			  ptrTempMatch=ptrOldTempMatch;
			  
			  while(ptrTempMatch->ptrNextMatch) {
				  ptrOldTempMatch=ptrTempMatch;
				  ptrTempMatch=ptrTempMatch->ptrNextMatch;
			  }
			  //If deletion was a failure (only match at specified beta row), then return a failure.
			  if(!ptrTempMatch->Delete_MN(ptrOldTempMatch)) return 0; 
		  }
     }
}
//Truncation was a success.
return 1;
}

//Routine used to clear the Alpha_Loc and ptrCol_Loc arrays so that the match structure can be re-used for a subsequent search of a repeated subfragment occurrence
//without having to recreate the entire Match Structure.
void
MatchStructure::Initialize_Arrays(unsigned _int32 intNum_Beta_Atoms) {

	unsigned _int32 i;

	if(ptrCol_Loc) {
		//Initialize arrays
		for(i=0;i<intNum_Beta_Atoms;i++) {
			ptrCol_Loc[i]=0;
		}
	}

	if(ptrAlpha_Loc) {
		for(i=0;i<ptrAlpha_Loc->intArray_Length;i++) {
			ptrAlpha_Loc->ptrQuery_Loc[i]=0;
		}
	}
}

//Creates a new match element node for a match element list.
void
MatchStructure::New_MN(MatchElement *&ptrTemp3,unsigned _int32 intDepth,unsigned _int32 intQuery_Loc){

	//Add to existing match element.
	if(ptrTemp3) {
		ptrTemp3->Add_MN(intQuery_Loc);
		ptrTemp3=ptrTemp3->ptrNextMatch;
	}
	//Else create at pointer array base.
	else {
		ptrMS_Spine[intDepth]=new MatchElement;
		ptrMS_Spine[intDepth]->intMatch=intQuery_Loc;
		ptrTemp3=ptrMS_Spine[intDepth];
	}

}

//Routine used to clear data structures in MatchStructure because destructors cannot be passed parameters.
void
MatchStructure::Clear_MS(unsigned _int32 intNum_Beta_Atoms) {

	unsigned _int32 i;
	unsigned _int32 j;

	//Delete diagonal storage structures.
	if(ptrRefine_Rack) {
		for(i=0;i<intNum_Beta_Atoms-1;i++) {
			for(j=0;j<=i;j++) {
				if(ptrRefine_Rack[i][j]) delete ptrRefine_Rack[i][j];
			}
			delete[] ptrRefine_Rack[i];
			delete[] ptrRefine_Top[i];
		}
		ptrRefine_Rack=0;
		ptrRefine_Top=0;
	}

	//Delete arrays.
	for(i=0;i<intNum_Beta_Atoms;i++) {
		delete ptrMS_Spine[i];
	}
	if(ptrMS_Spine) delete[] ptrMS_Spine;
	ptrMS_Spine=0;

	if(ptrCol_Loc) delete[] ptrCol_Loc;
	ptrCol_Loc=0;

	if(ptrAlpha_Loc) delete ptrAlpha_Loc;
	ptrAlpha_Loc=0;
}

//Constructor
MatchStructure::MatchStructure(unsigned _int32 intNum_Beta_Atoms,unsigned _int32 intNum_Alpha_Atoms,unsigned _int8 intSearch_Type) {

	unsigned _int32 i;
	unsigned _int32 j;

	const unsigned _int8 constBitSize=32;		 //Constant used to declare size of bit field integer used in storing detections during enumerating search
	unsigned _int32 intArray_Size;

	//If search is not an enumerating (combinatorial) search
	if(intSearch_Type<2) intArray_Size=intNum_Beta_Atoms;
	else intArray_Size=(intNum_Alpha_Atoms-1)/constBitSize+1;

	ptrAlpha_Loc=new SubFragLoc(intArray_Size);

	ptrMS_Spine=new MatchElement* [intNum_Beta_Atoms];
	ptrCol_Loc=new MatchElement* [intNum_Beta_Atoms];
	
	//Initialize arrays
	for(i=0;i<intNum_Beta_Atoms;i++) {
		ptrMS_Spine[i]=0;
		ptrCol_Loc[i]=0;
	}

	//Initialize diagonal storage structures
	ptrRefine_Rack=new MatchElement**[intNum_Beta_Atoms];
	ptrRefine_Top=new MatchElement**[intNum_Beta_Atoms];
	for(i=0;i<intNum_Beta_Atoms-1;i++) {
		ptrRefine_Rack[i]=new MatchElement*[i+1];
		ptrRefine_Top[i]=new MatchElement*[i+1];
	}
	for(i=0;i<intNum_Beta_Atoms-1;i++) {
		for(j=0;j<=i;j++) {
			ptrRefine_Rack[i][j]=0;
			ptrRefine_Top[i][j]=0;
		}
	}
}

MatchStructure::~MatchStructure() {

}

