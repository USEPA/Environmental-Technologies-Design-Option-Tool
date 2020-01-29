///////////////////////////////////////////////////////////////
//	This is the function code accompanying the 
//	header file SDA_1.h for MOSDAP(c) v2.0
//
//	Copyright (c) 1998. John W. Raymond.  All rights reserved.
//
///////////////////////////////////////////////////////////////

#include "SDA_1.H"


void AtomBond::AddConnection(unsigned _int8 intBondValue,unsigned _int32 intPreviousAtom,unsigned _int8 intNotBond) {

     ptrNextBond=new AtomBond;
     ptrNextBond->Bond_ID.Bond_Type=intBondValue;
	 if(intNotBond) ptrNextBond->Bond_ID.Not_Bond=1;
     ptrNextBond->intAttachedAtom=intPreviousAtom;
}

void SearchID::Reset_ID() {
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

void Atom::AddConnection(unsigned _int8 intBondValue,unsigned _int32 intPreviousAtom,unsigned _int8 intNotBond) {

     ptrNextBond=new AtomBond;
     ptrNextBond->Bond_ID.Bond_Type=intBondValue;
	 if(intNotBond) ptrNextBond->Bond_ID.Not_Bond=1;
     ptrNextBond->intAttachedAtom=intPreviousAtom;
}


//This routine fills the temporary quantity bit field with the contents of the current alpha graph quantity bit field.  
//This is used in non-truncating searches.
void 
QK_BF::Copy_QK(QK_BF &Tmp_QK){

     //Aliphatic Carbon
     Tmp_QK.Num_C=Num_C;
     //Aliphatic Oxygen
     Tmp_QK.Num_O=Num_O;
     //Aliphatic Nitrogen
     Tmp_QK.Num_N=Num_N;
     //Aliphatic Sulfur
     Tmp_QK.Num_S=Num_S;
     //Phosphorous
     Tmp_QK.Num_P=Num_P;
     //Bromine
     Tmp_QK.Num_Br=Num_Br;
	 //Fluorine
     Tmp_QK.Num_F=Num_F;
	 //Chlorine 
     Tmp_QK.Num_Cl=Num_Cl;
	 //Iodine
     Tmp_QK.Num_I=Num_I;
	 //Miscellaneous atoms (i.e., Si,Se,B,etc.)
     Tmp_QK.Num_Etc=Num_Etc;
	 //Double bonds
	 Tmp_QK.Num_DB=Num_DB;
	 //Triple bond
     Tmp_QK.Num_TB=Num_TB;
	 //Double bonded oxygens
	 Tmp_QK.Num_CDBO=Num_CDBO;
	 //Double bonded carbons
	 Tmp_QK.Num_CDBC=Num_CDBC;
	 //Triple bonded nitrogens
	 Tmp_QK.Num_CTBN=Num_CTBN;
	 //Aromatic carbon
     Tmp_QK.Num_AC=Num_AC;
     //Aromatic oxygen
     Tmp_QK.Num_AO=Num_AO;
     //Aromatic sulfur
     Tmp_QK.Num_AS=Num_AS;
     //Aromatic nitrogen
     Tmp_QK.Num_AN=Num_AN;
     //Number of rings
     Tmp_QK.Num_Rings=Num_Rings;
}

//Routine to reset the quantity keys to zero.
void 
QK_BF::Reset_QK(){

     //Aliphatic Carbon
     Num_C=0;
     //Aliphatic Oxygen
     Num_O=0;
     //Aliphatic Nitrogen
     Num_N=0;
     //Aliphatic Sulfur
     Num_S=0;
     //Phosphorous
     Num_P=0;
     //Bromine
     Num_Br=0;
	 //Fluorine
     Num_F=0;
	 //Chlorine 
     Num_Cl=0;
	 //Iodine
     Num_I=0;
	 //Miscellaneous atoms (i.e., Si,Se,B,etc.)
     Num_Etc=0;
	 //Double bonds
	 Num_DB=0;
	 //Triple bond
     Num_TB=0;
	 //Double bonded oxygens
	 Num_CDBO=0;
	 //Double bonded carbons
	 Num_CDBC=0;
	 //Triple bonded nitrogens
	 Num_CTBN=0;
	 //Aromatic carbon
     Num_AC=0;
	 //Aromatic oxygen
     Num_AO=0;
     //Aromatic sulfur
     Num_AS=0;
     //Aromatic nitrogen
     Num_AN=0;
	 //Number of rings
     Num_Rings=0;
}

//This routine decrements the quantity bit screen.
void
QK_BF::Truncate_QK(SubFragLoc *ptrAlpha_Loc,Atom *ptrAlpha_Molecule,Atom *ptrBeta_Molecule,unsigned _int32 intNum_Beta_Atoms) {

unsigned _int32 intBetaCntr;	//Loop counter
AtomBond *ptrTempBond;			//Temporary bond class pointer

for(intBetaCntr=0;intBetaCntr<intNum_Beta_Atoms;intBetaCntr++) {

     //Make sure detected atom is not a decharacterized atom
     if(!ptrBeta_Molecule[intBetaCntr].Search_ID.Dechar_Atom) {
          
          //Update the intQuantKey accordingly to reflect the subfragment detections
          switch(ptrAlpha_Molecule[ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Atom_ID.Atom_Type) {

          //Carbon
          case 6:
               //Aromatic carbon
               if(ptrAlpha_Molecule[ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Search_ID.Ring_Type == (1<<1)){ 
                    if(Num_AC) Num_AC--;
               }
               //Aliphatic ring or non-ring carbon
               else {
				   if(Num_CDBC && ptrAlpha_Molecule[ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Atom_ID.Bond_Types & (1<<1)) Num_CDBC--;
				   if(Num_C) Num_C--;
               }
               //Account for triple/double bonded pair keys incorporating carbon
			   ptrTempBond=ptrBeta_Molecule[intBetaCntr].ptrNextBond;
			   while(ptrTempBond) {
					//Only scan forward in the subfragment
				   if(ptrTempBond->intAttachedAtom>intBetaCntr) {
					   //Make sure bond is a double bond
					   if(ptrTempBond->Bond_ID.Bond_Type==2) {
							//Check attached atom type.
							switch(ptrBeta_Molecule[ptrTempBond->intAttachedAtom].Atom_ID.Atom_Type) {
							//Carbon
							case 6:
								if(Num_CDBC) Num_CDBC--;
								break;
							//Oxygen
							case 8:
								if(Num_CDBO) Num_CDBO--;
								break;
							}
							break;
					   }
					   //Make sure bond is a triple bond.
					   else if(ptrTempBond->Bond_ID.Bond_Type==3) {
							//Check attached atom type.
							switch(ptrBeta_Molecule[ptrTempBond->intAttachedAtom].Atom_ID.Atom_Type) {
							//Nitrogen
							case 7:
								if(Num_CTBN) Num_CTBN--;
								break;
							}
					   }
				   }
				   ptrTempBond=ptrTempBond->ptrNextBond;
			   }			   
			   break;
          //Oxygen
          case 8:
               //Aromatic oxygen
               if(ptrAlpha_Molecule[ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Search_ID.Ring_Type == (1<<1)){ 
                    if(Num_AO) Num_AO--;
               }
               //Aliphatic ring or non-ring oxygen
               else {
                   if(Num_CDBO && ptrAlpha_Molecule[ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Atom_ID.Bond_Types & (1<<1)) Num_CDBO--;
				   if(Num_O) Num_O--;
               }
			   //Account for double bonded pair keys incorporating oxygen.
			   ptrTempBond=ptrBeta_Molecule[intBetaCntr].ptrNextBond;
			   while(ptrTempBond) {
					//Only scan forward in the subfragment
				   if(ptrTempBond->intAttachedAtom>intBetaCntr) {
					   //Make sure bond is a double bond
					   if(ptrTempBond->Bond_ID.Bond_Type==2) {
							//Check attached atom type.
							switch(ptrBeta_Molecule[ptrTempBond->intAttachedAtom].Atom_ID.Atom_Type) {
							//Carbon
							case 6:
								if(Num_CDBO) Num_CDBO--;
								break;
							}
							break;
					   }
				   }
				   ptrTempBond=ptrTempBond->ptrNextBond;
			   }
               break;
          //Nitrogen
          case 7:
               //Aromatic nitrogen
               if(ptrAlpha_Molecule[ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Search_ID.Ring_Type == (1<<1)){ 
                    if(Num_AN) Num_AN--;
               }
               //Aliphatic ring or non-ring nitrogen
               else if(Num_N) Num_N--;
               
			   //Account for triple bonded pair keys incorporating carbon.
			   ptrTempBond=ptrBeta_Molecule[intBetaCntr].ptrNextBond;
			   while(ptrTempBond) {
					//Only scan forward in the subfragment
				   if(ptrTempBond->intAttachedAtom>intBetaCntr) {
					   //Make sure bond is a double bond
					   if(ptrTempBond->Bond_ID.Bond_Type==3) {
							//Check attached atom type.
							switch(ptrBeta_Molecule[ptrTempBond->intAttachedAtom].Atom_ID.Atom_Type) {
							//Carbon
							case 6:
								if(Num_CTBN) Num_CTBN--;
								break;
							}
							break;
					   }
				   }
				   ptrTempBond=ptrTempBond->ptrNextBond;
			   }
               break;
          //Sulphur
          case 16:
               //Aromatic sulfur
               if(ptrAlpha_Molecule[ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Search_ID.Ring_Type == (1<<1)){ 
                    if(Num_AS) Num_AS--;
               }
               //Aliphatic ring or non-ring carbon
               else {
                    if(Num_S) Num_S--;
               }
               break;
          //Fluorine
          case 9:
               if(Num_F) Num_F--;
               break;
          //Chlorine
          case 17:
               if(Num_Cl) Num_Cl--;
               break;
          //Bromine
          case 35:
               if(Num_Br) Num_Br--;
               break;
          //Iodine
          case 53:
               if(Num_I) Num_I--;
               break;
          //Phosphorous
          case 15:
               if(Num_P) Num_P--;
               break;
          //Selenium
          case 34:
               if(Num_Etc) Num_Etc--;
               break;         
          //Silicon
          case 14:
               if(Num_Etc) Num_Etc--;
               break;
          default:
               if(Num_Etc) Num_Etc--;
               break;
          }
     }
}
}

//  This routine performs a simple Quick Sort algorithm to sort a portion of two _int32 arrays of integers bounded by intLower_Bounds and intUpper_Bounds in ascending order based on an array specified as a "key" vector.
void 
ChemComboID::Bounded_QuickSort(ComboLoc **Key_Vector,unsigned _int32 *Slave_Vector,ComboLoc *ptrMax_Bound,unsigned _int32 intLower_Bounds,unsigned _int32 intLow,unsigned _int32 intUpper_Bounds,unsigned _int32 intHigh) {

	unsigned _int8 intFlag;
	unsigned _int8 intFlag2=0;
	unsigned _int32 i;
	unsigned _int32 j;
	_int32 intKey;
	_int32 intTemp_Slave;
	ComboLoc *ptrTemp_Key;
	ComboLoc *ptrTemp_Dummy=0;
	const _int32 intZero=0;
	
	//Perform Sort.
	if(intLow<intHigh) {

		//Store real value of array location following the sort segment's upper boundary, and
		//replace with a dummy large integer value.
		if(!ptrMax_Bound) {
			intFlag2=1;
			ptrMax_Bound=new ComboLoc;
			ptrMax_Bound->ptrChemID=new ChemSeqID;
			ptrMax_Bound->ptrChemID->intChemEntryID=~intZero;

			ptrTemp_Dummy=Key_Vector[intUpper_Bounds+1];
			Key_Vector[intUpper_Bounds+1]=ptrMax_Bound;
		}
		
		intFlag=1;
		i=intLow;
		j=intHigh+1;

		intKey=Key_Vector[intLow]->ptrChemID->intChemEntryID;

		while(intFlag) {

			i++;

			while(Key_Vector[i]->ptrChemID->intChemEntryID<intKey && i<=intHigh) {
				i++;
			}

			j--;

			while(Key_Vector[j]->ptrChemID->intChemEntryID> intKey && j>=intLow) {
				j--;
			}

			if(i<j) {
				//Interchange elements.
				ptrTemp_Key=Key_Vector[j];
				Key_Vector[j]=Key_Vector[i];
				Key_Vector[i]=ptrTemp_Key;

				intTemp_Slave=Slave_Vector[j];
				Slave_Vector[j]=Slave_Vector[i];
				Slave_Vector[i]=intTemp_Slave;
			}
			else intFlag=0;
		}
		
		//Interchange elements.
		ptrTemp_Key=Key_Vector[j];
		Key_Vector[j]=Key_Vector[intLow];
		Key_Vector[intLow]=ptrTemp_Key;

		intTemp_Slave=Slave_Vector[j];
		Slave_Vector[j]=Slave_Vector[intLow];
		Slave_Vector[intLow]=intTemp_Slave;

		//Recursively call Quick Sort routine.
		if(j>intLower_Bounds) Bounded_QuickSort(Key_Vector,Slave_Vector,ptrMax_Bound,intLower_Bounds,intLow,intUpper_Bounds,j-1);
		if(j<intUpper_Bounds) Bounded_QuickSort(Key_Vector,Slave_Vector,ptrMax_Bound,intLower_Bounds,j+1,intUpper_Bounds,intHigh);

		//Replace dummy array location with real value.
		if(intFlag2) {
			Key_Vector[intUpper_Bounds+1]=ptrTemp_Dummy;
		
			if(ptrMax_Bound) {
				delete ptrMax_Bound->ptrChemID;
				ptrMax_Bound->ptrChemID=0;
				delete ptrMax_Bound;
				ptrMax_Bound=0;
			}
		}
	}
}

//  This algorithm checks a current grouping cover to determine whether it is a degenerate occurrence.
unsigned _int8
ChemComboID::Check_Degeneracy(ComboLoc **ptrFragID,unsigned _int32 *ptrBitCntr,unsigned _int32 &intDetectionCntr,unsigned _int32 &intOldDemarcation,unsigned _int32 intNum_Groups[30],unsigned _int32 &intGroup_Cntr) {

	ComboLoc **ptrTemp_FragID=0;		//Temporary pointer used to create internal accounting array of detected subfragments.	
	ComboLoc *ptrJunk_Pointer=0;		//Junk pointer used to fill a parameter that is necessary only when Quick Sort is called recursively.
	unsigned _int32 *ptrTemp_BitCntr=0;	//Temporary pointer used to create internal accounting array to store vertex "level" of subfragment
	unsigned _int32 intNewDemarcation;
	unsigned _int32 intTemp_Cntr=0;		//Temporary counter variable for current cover.
	unsigned _int32 intTemp_GroupCntr;	//Temporary counter to account for each cover as a single entity.

	unsigned _int32 i;					//Junk counter variable.
	unsigned _int32 j;					//Junk counter variable.

	//Copy current subfragment order in current grouping to temporary arrays.
	ptrTemp_FragID=new ComboLoc*[intNum_Groups[intGroup_Cntr]];
	ptrTemp_BitCntr=new unsigned _int32[intNum_Groups[intGroup_Cntr]];
	intNewDemarcation=intDetectionCntr-1;

	for(j=intOldDemarcation;j<=intNewDemarcation;j++) {
		ptrTemp_FragID[intTemp_Cntr]=ptrFragID[j];
		ptrTemp_BitCntr[intTemp_Cntr]=ptrBitCntr[j];
		intTemp_Cntr++;
	}

	//Sort current subfragment grouping for degenerate comparison.
	Bounded_QuickSort(ptrFragID,ptrBitCntr,ptrJunk_Pointer,intOldDemarcation,intOldDemarcation,intNewDemarcation,intNewDemarcation);

	//Compare current grouping to previous groupings to ascertain whether it is degenerate.
	if(intOldDemarcation) {
		i=0;
		intTemp_GroupCntr=0;
		
		while(i!=(intOldDemarcation-1)) {
			
			if(ptrFragID[i]==0) i++;

			if(intNum_Groups[intTemp_GroupCntr]==(intDetectionCntr-intOldDemarcation)) {

				j=intOldDemarcation;
				while(ptrFragID[i]>0) {
					if(ptrFragID[i]->ptrChemID->intChemEntryID != ptrFragID[j]->ptrChemID->intChemEntryID) break;
					i++;
					j++;
				}

				//Cover is degenerate.
				if(ptrFragID[i]==0) {
			
					//Copy all subfragments in unsorted order except for the last one over the 
					//degenerate grouping to continue the cover detection.
					intDetectionCntr--;
					ptrFragID[intDetectionCntr]=0;
					ptrBitCntr[intDetectionCntr]=0;
					
					intTemp_Cntr-=2;
					for(j=(intDetectionCntr-1);j>=intOldDemarcation;j--) {
						ptrFragID[j]=ptrTemp_FragID[intTemp_Cntr];
						ptrBitCntr[j]=ptrTemp_BitCntr[intTemp_Cntr];
						if(intTemp_Cntr) intTemp_Cntr--;
					}

					delete[] ptrTemp_FragID;
					delete[] ptrTemp_BitCntr;
					return 0;
				}
			}
			intTemp_GroupCntr++;

			//Skip to next grouping
			while(ptrFragID[i]!=0) {
				i++;
			}
		}
	}
	
	//Not a degenerate cover.
	//Copy all but last subfragment in the unsorted order to new grouping set.
	intDetectionCntr++;
	for(j=0;j<(intTemp_Cntr-1);j++) {
		ptrFragID[intDetectionCntr]=ptrTemp_FragID[j];
		ptrBitCntr[intDetectionCntr]=ptrTemp_BitCntr[j];
		intDetectionCntr++;
	}
	intOldDemarcation=intNewDemarcation+2;

	delete[] ptrTemp_FragID;
	delete[] ptrTemp_BitCntr;
	return 1;
}

//Bitwise OR's the subfragment detection onto the exact cover checking structure.
void
ChemComboID::Fill_EC_Check(ComboLoc *ptrTempComboLoc) {

	unsigned _int32 i;

	for(i=0;i<intEC_Array_Length;i++) {
	   ptrEC_Check[i] |= (ptrTempComboLoc->ptrQuery_Loc[i]);
    }
                                             
}

//Initializes the EC_Check array used to determine whether the possibility of an exact cover of the query exists from the pool of subfragment detections to
//determine whether it is necessary to proceed to the exact cover algorithm.
void
ChemComboID::Initialize_EC_Check() {

	unsigned _int32 i;
	const unsigned _int8 constCB_Size=32;

	intEC_Array_Length=(intNumberAtoms-1) / constCB_Size + 1;
    ptrEC_Check=new unsigned _int32[intEC_Array_Length];
    
	for(i=0;i<intEC_Array_Length;i++) {
	    ptrEC_Check[i]=0;
    }
}

//This routine compares the EC_Check array with a completely filled (bitwise) array to determine whether all atoms have been detected in at least one
//subfragment before proceeding on with the exact cover algorithm.
unsigned _int8
ChemComboID::Compare_EC_Check() {

	unsigned _int32 i;
	const unsigned _int32 constZero=0;
	const unsigned _int8 constCB_Size=32;	//Bit size of integers used in arrays.
	unsigned _int8 intResid_Length;

	intResid_Length=(intNumberAtoms-1) % constCB_Size;
	
	//Check if the cummulative collection of detections completely covers the alpha graph before proceeding with set cover algorithm.
    //Check all but last (potentially not completely filled) integer bit field.
    for(i=0;i<(intEC_Array_Length-1);i++) {
		if(ptrEC_Check[i] ^ (~constZero)) return 0;
	}
    //Check last (completely filled) bit field.
    if(intResid_Length == (constCB_Size-1)) {
		if(ptrEC_Check[intEC_Array_Length-1] ^ (~constZero)) return 0;
    }
    //Check last (not completely filled) bit field (array integer location).
    else {
		if(ptrEC_Check[intEC_Array_Length-1] ^ ((1<<(intResid_Length+1))-1)) return 0;
    }
	return 1;
}          

//This routine performs a list pocket sort of the subfragment list based on the occupancy key.
//   | 11 | 10 | 9 |  8    |   7    | 6 |   5    | 4 |   3    |   2  |  1   |    0    |
//   | remaining query entries | C | c | s,o,n | rings? | O | DB (=) | N | TB (#) | Cl,I | F,Br | Etc,S,P |
void 
SubFragList::ListBinSort(unsigned _int8 intNumPockets){

ChemSeqID **ptrLastPocket;
ChemSeqID *ptrTempNode;
_int8 i;
unsigned _int8 intPocketCntr;

ptrLastPocket=new ChemSeqID*[intNumPockets+1];
ptrBinSortLoc=new ChemSeqID*[intNumPockets+1];

//Initialize arrays
for(i=0;i<=intNumPockets;i++) {
     ptrLastPocket[i]=0;
     ptrBinSortLoc[i]=0;
}

ptrTempNode=ptrFirstIDNode;

//Distribute list structure into pockets
while(ptrTempNode) {

     intPocketCntr=0;

     //Loop through occupation key until an appropriate bucket is found
     while(intPocketCntr<intNumPockets) {

          //If key is present in subfragment or last (catch all) bin is encountered.
		  if((ptrTempNode->intOccupancyKey & (1<<intPocketCntr)) || (intPocketCntr==(intNumPockets-1))) {

               //If pocket is temporarily empty
               if(!ptrBinSortLoc[intPocketCntr]) {

                    ptrLastPocket[intPocketCntr]=ptrTempNode;
                    ptrBinSortLoc[intPocketCntr]=ptrTempNode;
               }
               //Else if pocket is not empty
               else {
                    ptrLastPocket[intPocketCntr]->ptrNextSubFrag=ptrTempNode;
                    ptrLastPocket[intPocketCntr]=ptrTempNode;
               }

               break;
          }
          intPocketCntr++;

     }
     ptrTempNode=ptrTempNode->ptrNextSubFrag;
     if(ptrLastPocket[intPocketCntr])ptrLastPocket[intPocketCntr]->ptrNextSubFrag=0;

}

//Recombine the pocket sorted list
intPocketCntr=0;

//Find first non-null location in bin array
while(!ptrBinSortLoc[intPocketCntr]) {
     
     intPocketCntr++;
}

ptrFirstIDNode=ptrBinSortLoc[intPocketCntr];

//Build list
for(intPocketCntr+=1;intPocketCntr<intNumPockets;intPocketCntr++) {

     ptrTempNode=ptrLastPocket[intPocketCntr-1];

     if(ptrLastPocket[intPocketCntr]) {
          ptrTempNode->ptrNextSubFrag=ptrBinSortLoc[intPocketCntr];
     }
     else {
          ptrLastPocket[intPocketCntr]=ptrTempNode;
     }
}

ptrTempNode=0;

//Fill null locations with pointer from following non-null depth
for(i=(intNumPockets-1);i>=0;i--) {

     if(ptrBinSortLoc[i]) {
          ptrTempNode=ptrBinSortLoc[i];
     }
     if(i<(intNumPockets-1) && !ptrBinSortLoc[i]) {
          ptrBinSortLoc[i]=ptrTempNode;
     }
}
delete[] ptrLastPocket;
}

//Read in subfragment list.
unsigned _int8 
SubFragList::Read_SF_List(char *strSubFragFile,_int8 intQueryType){

ChemSeqID *ptrTempNode;
unsigned _int16 intTempID;

//Create dynamic character buffer array to read in beta subfragments
char *ptrTempString;
ptrTempString=new char[610];

ptrFirstIDNode=new ChemSeqID;
ptrTempNode=ptrFirstIDNode;

//Open subfragment file
ifstream Input_File;
Input_File.open(strSubFragFile);

//Read through the subfragment file until end of file and store into the subfragment string list
while(!Input_File.eof()) {

     //Read subfragment string from file
     Input_File>>intTempID;
     
     //If end of file (eof) then break from file input loop
     if(Input_File.eof()) {
		 if(!ptrFirstIDNode) {
			 delete[] ptrTempString;
			 Input_File.close();
			 return 0;
		 }
		 delete[] ptrTempString;
		 Input_File.close();
		 break;
	 }

     //Create new ChemSeqID node for new chemical structure string
     if(intNumber_SF) {
          ptrTempNode->ptrNextSubFrag=new ChemSeqID;
          ptrTempNode=ptrTempNode->ptrNextSubFrag;
     }

     //Read subfragment string from file
     Input_File>>ptrTempString;
     
	 //Get length of the string and allocate memory for chemical syntax string.
	 ptrTempNode->chrChemSyntaxString=new char[strlen(ptrTempString)+1];
	 ptrTempNode->intChemEntryID=intTempID;
	 strcpy(ptrTempNode->chrChemSyntaxString,ptrTempString);

     //If "file search" option is selected, then fill key bit fields for future screening/sorting
     if(intQueryType) {
          ptrTempNode->SMILES_QK_Fill();
          //Call function to convert quantity key to occupancy key
          ptrTempNode->QK_to_OK();
          intAvg_Num_Atoms += ptrTempNode->intNumberAtoms;
     }

     //Increment the number of subfragments counter
     intNumber_SF++;
}

if(intNumber_SF) intAvg_Num_Atoms /= intNumber_SF;
else return 0;

return 1;
}