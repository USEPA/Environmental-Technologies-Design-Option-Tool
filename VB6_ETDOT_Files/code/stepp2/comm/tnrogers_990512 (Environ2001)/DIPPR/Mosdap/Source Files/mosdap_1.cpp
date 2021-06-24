///**********************************************************************************************************
//        MOlecular Structure DissAssembly Program (MOSDAP) 2.0
//
//**********************************************************************************************************
//
//        Copyright (c) by John W. Raymond, Jr., 1998
//        
//        All rights reserved. No part of this program in any form, compiled, uncompiled, or 
//        or otherwise, shall be reproduced or transmitted by any means, electronic, mechanical, physical or 
//        otherwise, without the express written consent of the author.  Usage of the program in a manner other
//        than that explicitly stated in the written consent of the author is prohibited.
//
//**********************************************************************************************************
//        This program performs a substructure search on a chemical structure represented by SMILES nomenclature.
//        The ASCII input substructure search file consists of chemical substructures represented by a customized, pseudo-
//        SMILES nomenclature.  The substructure search can be performed as a truncating search for group contribution
//        type correlations or a non-truncating search for pure substructure searching.  The program is coded to
//        be used either as a stand alone program or as an exportable library routine in a Windows dynamic link library
//
//**********************************************************************************************************
//        This is a construction version incorporating the new data types and molecular feature searching.
//**********************************************************************************************************

//Include customized graph structure header file.  Contains all classes used by the MOSDAP program.
#include "SDA_1.H"
#include "SDA_2.H"

#include <windows.h>
#include <iostream.h>
#include <io.h>
#include <strstrea.h>

const TEMP_MAX=512;

#define DLLAPI WINAPI

#if defined(DEBUG)
#define ASSERT(f)	\
	if(f)	\
	Null;	\
	else 
	assert(f)
#else
#define ASSERT(f) NULL
#endif

void ErrorHandler(long e);
DWORD HResultToErr(long e);

HINSTANCE hInst;

BOOL WINAPI DllMain(HINSTANCE hInstA, DWORD dwReason, LPVOID lpvReserved){

	switch(dwReason) {
		case DLL_PROCESS_ATTACH:
			hInst=hInstA;
			break;

		case DLL_THREAD_ATTACH:
		break;

		case DLL_THREAD_DETACH:
			break;

		case DLL_PROCESS_DETACH:
			hInst=0;
			break;
	}
	return TRUE;
}


extern "C" void _declspec(dllexport) _stdcall MOSDAP(char *strQuery,_int8 intQueryType,char *strSubFragFile,char *strOutputQueryFile,_int8 intSearchType,_int8 &intSearchResult,_int32 intSF_ID[100],_int32 intSF_Quant[100],_int32 intMF_ID[21],_int32 intMF_Quant[21]);

//This routine performs a simple atom quantity pre-screen of a master (alpha) and subfragment (beta) structure
//using a simple comparison of the quantities of certain atom and bond types (using the QK_BF class).
unsigned _int8 
QK_Screen(QK_BF &intAlphaQuantKey,QK_BF &intSF_QuantKey) {

     //Check to make sure that the subfragment (beta) structure does not contain more of a
     //certain atom or bond type than the master (alpha) structure.

     //Aliphatic Carbon
     if(intAlphaQuantKey.Num_C<intSF_QuantKey.Num_C) return 0;
     //Aliphatic Oxygen
     else if(intAlphaQuantKey.Num_O<intSF_QuantKey.Num_O) return 0;
     //Aliphatic Nitrogen
     else if(intAlphaQuantKey.Num_N<intSF_QuantKey.Num_N) return 0;
     //Aliphatic Sulfur
     else if(intAlphaQuantKey.Num_S<intSF_QuantKey.Num_S) return 0;
     //Double bonds
     else if(intAlphaQuantKey.Num_DB<intSF_QuantKey.Num_DB) return 0;
     //Aromatic carbon
     else if(intAlphaQuantKey.Num_AC<intSF_QuantKey.Num_AC) return 0;
     //Chlorine 
     else if(intAlphaQuantKey.Num_Cl<intSF_QuantKey.Num_Cl) return 0;
     //Fluorine
     else if(intAlphaQuantKey.Num_F<intSF_QuantKey.Num_F) return 0;
     //Bromine
     else if(intAlphaQuantKey.Num_Br<intSF_QuantKey.Num_Br) return 0;
     //Iodine
     else if(intAlphaQuantKey.Num_I<intSF_QuantKey.Num_I) return 0;
     //Triple bond
     else if(intAlphaQuantKey.Num_TB<intSF_QuantKey.Num_TB) return 0;
     //Carbon to double bonded oxygens
     else if(intAlphaQuantKey.Num_CDBO<intSF_QuantKey.Num_CDBO) return 0;
	 //Carbon to double bonded carbons
     else if(intAlphaQuantKey.Num_CDBC<intSF_QuantKey.Num_CDBC) return 0;
	 //Carbon to triple bonded nitrogens
     else if(intAlphaQuantKey.Num_CTBN<intSF_QuantKey.Num_CTBN) return 0;
	 //Miscellaneous atoms (i.e., Si,Se,B,etc.)
     else if(intAlphaQuantKey.Num_Etc<intSF_QuantKey.Num_Etc) return 0;
     //Aromatic oxygen
     else if(intAlphaQuantKey.Num_AO<intSF_QuantKey.Num_AO) return 0;
     //Aromatic sulfur
     else if(intAlphaQuantKey.Num_AS<intSF_QuantKey.Num_AS) return 0;
     //Aromatic nitrogen
     else if(intAlphaQuantKey.Num_AN<intSF_QuantKey.Num_AN) return 0;
     //Number of rings
     else if(intAlphaQuantKey.Num_Rings<intSF_QuantKey.Num_Rings) return 0;
     //Phosphorous
     else if(intAlphaQuantKey.Num_P<intSF_QuantKey.Num_P) return 0;
     //Passed pre-screen
     else return 1;
}

//This routine truncates (usually temporarily) the alpha graph so that the same atoms are not used in additional substructure searching (for use with
//sequential truncating searches). 
inline void
Truncate_Molecule(MatchStructure *ptrMS,ChemSeqID *ptrAlpha,ChemSeqID *ptrBeta,unsigned _int32 &intUsedAtoms) {

unsigned _int32     intBetaCntr;

for(intBetaCntr=0;intBetaCntr<ptrBeta->intNumberAtoms;intBetaCntr++) {

     //Set the "already used" flag in the alpha structure intAtomID bit field if it does not correspond to a decharacterized atom in the beta structure
     if(!ptrBeta->ptrMolecule->ptrAtom[intBetaCntr].Search_ID.Dechar_Atom) {
          
          //Increment the number of used atoms in alpha molecule
          intUsedAtoms++;

          //Set "already used" bit 
          ptrAlpha->ptrMolecule->ptrAtom[ptrMS->ptrAlpha_Loc->ptrQuery_Loc[intBetaCntr]].Search_ID.Used_Atom=1;
          
     }
}

}

//   The algorithm for this code is a modified version of the algorithm due to Kreher, 1997. It
//   has been modified to exclude recursion, allow dynamicly declared data objects, and facilitate
//   the use of pointers.
unsigned _int8
KreherExactCover(ChemComboID *ptrCombo_Node,_int32 SubfragID[100],_int32 SubfragQuant[100]) {

//CODE BLOCK to sort list of potential subfragments into pockets of decreasing lexicographical order
//and to fill the Change array pointing to each location where a bit change occurs.

ComboLoc **ptrLastPocket;               //Pointer vector storing pointer address pointing to the last element in a pocket
ComboLoc **ptrFragID;                   //Pointer used to create internal accounting array of detected subfragments
ComboLoc **ptrECChange;                 //Pointer for array used to store where "level" changes occur in subfragment list (from bin sort)
ComboLoc *ptrTemp1Node;                 //Junk temporary pointer.
ComboLoc *ptrTemp2Node;                 //Junk temporary pointer.
unsigned _int32 intNum_Groups[30]={0};	//Static array to store then quantity of subfragments in each non-degenerate cover. (Max=>30 unique covers).
unsigned _int32 *ptrBitCntr;			//Pointer used to create internal accounting array to store vertex "level" of subfragment
unsigned _int32 *ptrPartialGate;		//Array pointer used to construct the cummulative structure using detected subfragments
unsigned _int32 i;		                //Temporary counter variable.
unsigned _int32 j;                      //Temporary counter varialbe.
unsigned _int8 intJumpFlag;             //Flag to determine whether to use Change or Next element (0~Change, 1~Next)
unsigned _int32 intPocketCntr;          //Counter variable storing location of pocket.
unsigned _int32 intECChangePlace;		//ECChange counter variable
const unsigned _int8 intBitFieldSize=32;//Stores the size of the integers used as bitfields for storing subfragment locations.
unsigned _int32 intArrayCntr=0;			//Stores location in array structure representing subfragment locations of current bit.
unsigned _int32 intCurrentBitCntr=0;	//Stores bit location in integer bitfield specified by intArrayCntr.
unsigned _int32 intCurrentChangeCntr=0; //Counter for the ECChange[] used to end search at given level
unsigned _int32 intDetectionCntr=0;     //Current grouping subfragment detection counter
unsigned _int32 intGroup_Cntr=0;			//Counter to account for number of non-degenerate  (unique) covers.
unsigned _int32 intUnique_Detect;		//Counter used to tabulate number of unique subfragment detections for transferring results to export arrays.
unsigned _int32 intOldDemarcation;		//Stores previous demarcation (for previous grouping) in accounting arrays
unsigned _int32 intOldBitCntr;          //Stores old bit counter
unsigned _int32 intStop;				//Junk variable for loop boundary
unsigned _int32 intStart;               //Junk variable for loop boundary
unsigned _int32 intNextLevel;			//Variable used to determine where next non-zero occurrence in ECChange[~] array is located
unsigned _int8 intBackTrack;			//Variable flag used to specify whether to backtrack or not
unsigned _int16 intMax_Detect;			//Maximum number of detections in grouping array.

//Allocate accounting buffers for maximum number of subfragments involved in all covers.
if(ptrCombo_Node->intNumberAtoms<10) intMax_Detect=150;
else if(ptrCombo_Node->intNumberAtoms<15) intMax_Detect=300;
else intMax_Detect=500;

//Create partial gate structure (used to accumulate subfragments to determine if structure is completely covered)
ptrPartialGate=new unsigned _int32[ptrCombo_Node->intEC_Array_Length];
for(i=0;i<ptrCombo_Node->intEC_Array_Length;i++) {
     ptrPartialGate[i]=0;
}

//Create temporary detection array (max # of detections => inMax_Detect)
ptrFragID=new ComboLoc*[intMax_Detect];
ptrBitCntr=new unsigned _int32[intMax_Detect];
for(i=0;i<intMax_Detect;i++) {
     ptrFragID[i]=0;
     ptrBitCntr[i]=0;
}

//Create pointer vectors
ptrLastPocket=new ComboLoc*[ptrCombo_Node->intNumberAtoms+1];
ptrECChange=new ComboLoc*[ptrCombo_Node->intNumberAtoms+1];

//Initialize arrays
for(i=0;i<=ptrCombo_Node->intNumberAtoms;i++) {
     ptrLastPocket[i]=0;
     ptrECChange[i]=0;
}

ptrTemp1Node=ptrCombo_Node->ptrFirstComboLoc;

//Distribute list structures into pockets.
while(ptrTemp1Node){
     intPocketCntr=ptrCombo_Node->intNumberAtoms-1;
     
     //Get position of initial bit counters.
     intArrayCntr=intPocketCntr/intBitFieldSize;
     intCurrentBitCntr=intPocketCntr%intBitFieldSize;
     intECChangePlace=ptrCombo_Node->intNumberAtoms-intPocketCntr-1;

     //Loop through the subfragment bitfield until an appropriate pocket is found.
     while(1) {

          //Obtain a radix single bit digit (0 or 1).
          if(ptrTemp1Node->ptrQuery_Loc[intArrayCntr] & (1<<intCurrentBitCntr)) {
               
               //If pocket is temporarily empty
               if(!ptrECChange[intECChangePlace]) {
                    ptrLastPocket[intECChangePlace]=ptrTemp1Node;
                    ptrECChange[intECChangePlace]=ptrTemp1Node;
               }
               //Else if pocket is not empty
               else{
                    ptrLastPocket[intECChangePlace]->ptrNextComboLoc=ptrTemp1Node;
                    ptrLastPocket[intECChangePlace]=ptrTemp1Node;
               }
               break;
          }

          if(intPocketCntr) {
               intPocketCntr--;
               //Get position of new bit counters.
               intArrayCntr=intPocketCntr/intBitFieldSize;
               intCurrentBitCntr=intPocketCntr%intBitFieldSize;
               intECChangePlace=ptrCombo_Node->intNumberAtoms-intPocketCntr-1;
          }
          else break;
     }

     ptrTemp1Node=ptrTemp1Node->ptrNextComboLoc;
     ptrLastPocket[intECChangePlace]->ptrNextComboLoc=0;
}
intPocketCntr=0;

//Find first non-null location in bin array. This coding does not check for blank nodes because it
//is assumed that these were detected prior to the instancing of the vertex exact cover function.
while(!ptrECChange[intPocketCntr]) {
     intPocketCntr++;
}
ptrCombo_Node->ptrFirstComboLoc=ptrECChange[intPocketCntr];

//Now reconstruct the pocket sorted list in decreasing order
for(intPocketCntr += 1;intPocketCntr<ptrCombo_Node->intNumberAtoms;intPocketCntr++){

     ptrTemp1Node=ptrLastPocket[intPocketCntr-1];
     if(ptrLastPocket[intPocketCntr]) {
          ptrTemp1Node->ptrNextComboLoc=ptrECChange[intPocketCntr];
     }
     else {
          ptrLastPocket[intPocketCntr]=ptrTemp1Node;
     }
}

delete[] ptrLastPocket;
ptrLastPocket=0;

//CODE BLOCK to determine the position of the first subgraph that is disjoint from the query graph.
ptrTemp1Node=ptrCombo_Node->ptrFirstComboLoc;

//Loop through portion of list below current subfragment to find first disjoint occurrence
while(ptrTemp1Node) {

     ptrTemp2Node=ptrTemp1Node->ptrNextComboLoc;

     while(ptrTemp2Node) {

          for(i=0;i<ptrCombo_Node->intEC_Array_Length;i++) {
               
               //If subfragments are overlapping, then skip to next subfragment
               if(ptrTemp1Node->ptrQuery_Loc[i] & ptrTemp2Node->ptrQuery_Loc[i]) {
                    ptrTemp2Node=ptrTemp2Node->ptrNextComboLoc;
                    break;
               }
          }
          //If subfragments are disjoint, break loop
          if(i==ptrCombo_Node->intEC_Array_Length) break;
     }
     if(ptrTemp2Node) ptrTemp1Node->ptrNextDisjoint=ptrTemp2Node;
     else ptrTemp1Node->ptrNextDisjoint=0;

     ptrTemp1Node=ptrTemp1Node->ptrNextComboLoc;
}

//CODE BLOCK to perform the Exact Cover backtrack to determine all possible combinations of subfragments that
//totally cover the query molecule.
intCurrentChangeCntr=0;       //Counter for ECChange[] array used to end search level
intDetectionCntr=0;           //Current grouping subfragment counter
intCurrentBitCntr=ptrCombo_Node->intNumberAtoms-1;   //Bit (alpha atom) location where first "not used" location in alpha molecule
intOldDemarcation=0;          //Stores location denoting previous grouping demarcation
intBackTrack=0;                    //Backtrack flag (0~ no backtrack, 1~ backtrack)
intNextLevel=1;                    //Counter used to specify next non-null element in the ECChange array

ptrTemp2Node=ptrCombo_Node->ptrFirstComboLoc;

while(1) {

     //Loop to step thru current subfrag level
     while(1){

          //Loop through entire alpha structure array
          for(j=0;j<ptrCombo_Node->intEC_Array_Length;j++) {
               //If subfrag overlaps gate structure
               if(ptrTemp2Node->ptrQuery_Loc[j] & ptrPartialGate[j]) {
                    break;
               }
          }

          //If subfrag is disjoint
          if(j==ptrCombo_Node->intEC_Array_Length) {

               //Subfragment detection vector code
               ptrFragID[intDetectionCntr]=ptrTemp2Node;
               ptrBitCntr[intDetectionCntr]=intCurrentBitCntr;
               intOldBitCntr=intCurrentBitCntr;
               
			   intDetectionCntr++;
			                  
			   //Append detected subfragment to cummulative subfragment structure (gate structure)
               for(j=0;j<ptrCombo_Node->intEC_Array_Length;j++) {
                    ptrPartialGate[j] |= ptrTemp2Node->ptrQuery_Loc[j];
               }

               //Check if gate structure is a complete replica of the query structure
               for(j=0;j<ptrCombo_Node->intEC_Array_Length;j++) {
                    if(ptrPartialGate[j] ^ ptrCombo_Node->ptrEC_Check[j]) break;
               }
               //If partial gate is complete, then add grouping demarcation to detection vectors
               if(j==ptrCombo_Node->intEC_Array_Length) {
                    //Add demarcation
                    ptrFragID[intDetectionCntr]=0;

					//Temporarily store number of subfragments in current grouping in accounting array.
					intNum_Groups[intGroup_Cntr]=(intDetectionCntr-intOldDemarcation);

					//Check if cover is degenerate and update group (cover) accounting variables.
					if(ptrCombo_Node->Check_Degeneracy(ptrFragID,ptrBitCntr,intDetectionCntr,intOldDemarcation,intNum_Groups,intGroup_Cntr)) {
						intGroup_Cntr++;
					}
					else intNum_Groups[intGroup_Cntr]=0;
										
					//Extract last subfragment from cummulative accounting array so that all possible combos of subfragments can be examined
                    for(j=0;j<ptrCombo_Node->intEC_Array_Length;j++) {
                         ptrPartialGate[j] ^= ptrTemp2Node->ptrQuery_Loc[j];
                    }

                    ptrTemp2Node=ptrTemp2Node->ptrNextComboLoc;

                    //Find next non-null occurrence in ptrECChange[~] array to serve as the terminus for current search "level
                    while(ptrECChange[intCurrentChangeCntr + intNextLevel] == 0 && (intCurrentChangeCntr + intNextLevel < ptrCombo_Node->intNumberAtoms)) {
                         intNextLevel++;
                    }

                    //Check to see if end of subfrag level has been encountered 
                    if(ptrTemp2Node == ptrECChange[intCurrentChangeCntr+intNextLevel]) {
                         intBackTrack=1;
                    }
                              
                    break;
               }

               //Find first open vertex position in partial subfragment gate structure 
               while(ptrPartialGate[intCurrentBitCntr/intBitFieldSize] & (1<<(intCurrentBitCntr%intBitFieldSize))){
                    intCurrentBitCntr--;
               }

               //Jump to next possible location in subfragment list to continue search
               intCurrentChangeCntr=ptrCombo_Node->intNumberAtoms-1-intCurrentBitCntr;

               //If a Next element and a Change element exist, then find the one farthest down the subfag list
               if(ptrTemp2Node->ptrNextDisjoint && ptrECChange[intCurrentChangeCntr]) {

                    intStart=intOldBitCntr/intBitFieldSize;
                    intStop=intCurrentBitCntr/intBitFieldSize;
                    intJumpFlag=1;
                    intNextLevel=1;

                    //Check to see if Change element is further down the list. If so, choose Change as the next subfrag.
                    for(j=intStart;j<=intStop;j--){
                         if(ptrTemp2Node->ptrNextDisjoint->ptrQuery_Loc[j] > ptrECChange[intCurrentChangeCntr]->ptrQuery_Loc[j] && !(ptrTemp2Node->ptrNextDisjoint->ptrQuery_Loc[j] & (1<<intCurrentBitCntr))) {
                              ptrTemp2Node=ptrECChange[intCurrentChangeCntr];
                              intJumpFlag=0;
                              break;
                         }
                    }
                    //Check to see if Next is located in the current Change level. If so, choose Next as the next subfrag.
                    if(intJumpFlag && (ptrTemp2Node->ptrNextDisjoint->ptrQuery_Loc[intStop] & (1<<intCurrentBitCntr))) ptrTemp2Node=ptrTemp2Node->ptrNextDisjoint;
                    //If none of the above, then backtrack because a completely "filled" structure cannot result
                    else if(intJumpFlag) {
                         intBackTrack=1;
                         break;
                    }
                    
               }
               //Else current subfragment cannot result in a completely "filled" structure
               else {
                    intBackTrack=1;
                    break;
               }
          }

          //Else subfragment is not disjoint. Go to next subfragment.
          else ptrTemp2Node=ptrTemp2Node->ptrNextComboLoc;

          //Find next non-null occurrence in ptrECChange[~] array to serve as the terminus for current search "level
          while(ptrECChange[intCurrentChangeCntr + intNextLevel] == 0 && (intCurrentChangeCntr + intNextLevel < ptrCombo_Node->intNumberAtoms)) {
               intNextLevel++;
          }

          //Check to see if end of subfrag level has been encountered 
          if(ptrTemp2Node == ptrECChange[intCurrentChangeCntr+intNextLevel]) {
               intBackTrack=1;
               break;
          }

     } //end current subfrag while loop

     //BACKTRACK A SUBFRAGMENT CODE BLOCK
     //If no disjoint subfragment in the current span is found
     while(intBackTrack)  {

          //Clear current invalid subfragment
          ptrFragID[intDetectionCntr]=0;
          ptrBitCntr[intDetectionCntr]=0;
          
          //End of set cover algorithm (all groupings investigated)
          if(intDetectionCntr ==0) {
			   //Clean up code for dynamically created arrays.
               delete[] ptrECChange;
               delete[] ptrPartialGate;
               delete[] ptrFragID;
               delete[] ptrBitCntr;
			   return 0;
		  }
		  else if(ptrFragID[intDetectionCntr - 1]==0) {

               //Code to put results in appropriate export vectors for export
               intOldDemarcation=0;
			   intUnique_Detect=0;
               for(i=0;i<intDetectionCntr-1;i++) {

                    for(j=intOldDemarcation;j<intUnique_Detect;j++) {
                         if(ptrFragID[i]) {
                              //Check if current subfragment is a duplicate of a previous subfragment in the current grouping. If so then increment the quantity.
							 if(ptrFragID[i]->ptrChemID->intChemEntryID==SubfragID[j]) {
                                   SubfragQuant[j]++;
                                   break;
                              }
                         }
                         //If ptrFragID[~] is a demarcation, then insert demarcation into the export vectors. 
                         else{
                              SubfragID[intUnique_Detect]=-1;
                              SubfragQuant[intUnique_Detect]=-1;
                              intUnique_Detect++;
							  intOldDemarcation=intUnique_Detect;
                              break;
                         }
                    }
                    //If no duplicate subfragment was detected, then simply append to the export vectors.
                    if(j==intUnique_Detect && ptrFragID[i]) {
                         SubfragID[intUnique_Detect]=ptrFragID[i]->ptrChemID->intChemEntryID;
                         SubfragQuant[intUnique_Detect]++;
						 intUnique_Detect++;
                    }
               }

               //Clean up code for dynamically created arrays.
               delete[] ptrECChange;
               delete[] ptrPartialGate;
               delete[] ptrFragID;
               delete[] ptrBitCntr;
               
               //Return status of covering process
               return 1;
          }
		  
          //Backtrack a subfragment
          intDetectionCntr--;
          ptrTemp2Node=ptrFragID[intDetectionCntr];
          
          //Update bit and change counters
          intCurrentBitCntr=ptrBitCntr[intDetectionCntr];
          intCurrentChangeCntr=ptrCombo_Node->intNumberAtoms-1-intCurrentBitCntr;

          //Extract detected subfragment from cummulative structure
          for(j=0;j<ptrCombo_Node->intEC_Array_Length;j++) {
               ptrPartialGate[j] ^= ptrTemp2Node->ptrQuery_Loc[j];
          }

          //Find next non-null occurrence in ptrECChange[~] array to serve as the terminus for current search "level
          intNextLevel=1;
          while(ptrECChange[intCurrentChangeCntr + intNextLevel] == 0 && (intCurrentChangeCntr + intNextLevel < ptrCombo_Node->intNumberAtoms)) {
               intNextLevel++;
          }

          //If backtracked subfragment is not the final subfrag in the current bit counter level, then exit backtrack.
          if(ptrTemp2Node->ptrNextComboLoc != ptrECChange[intCurrentChangeCntr + intNextLevel]) {

               ptrTemp2Node=ptrTemp2Node->ptrNextComboLoc;
               intBackTrack=0;
          }
     }

} //end while(1)

}

//Primitive depth-first search routine used for the perfect maximum bipartite matching of the augmented atom complex in the refinenment procedure.
unsigned _int8 
DFS_Max_Bip(unsigned _int8 intAlphaMatchCnt,unsigned _int8 intBetaMatchCnt,unsigned _int8 intMaxBipMatch[6][6]) {

const unsigned _int8 intNum_Neighbors=6;		//Number of neighbors in augmented atom complex.
unsigned _int8 i;								//Temporary counter.
unsigned _int8 j;								//Temporary counter.
unsigned _int8 intColCnt;						//intMaxBipMatch[~][~] Column counter
unsigned _int8 intMatchDepth;					//intMaxBipMatch[~][~] Depth counter
unsigned _int8 intColsUsed[intNum_Neighbors];	//Vector that stores which columns have already been used
_int8 intColatDUsed[intNum_Neighbors];			//Vector that stores which column at specified depth has been used


//Initialize the arrays for the matching to zero
for(i=0;i <intNum_Neighbors;i++) {
	intColsUsed[i]=0;
    intColatDUsed[i]=-1;
}
                              
//Resort to a primitive backtrack procedure to determine if there is a 1:1 correspondence between alpha and beta
intMatchDepth=0;
intColCnt=0;

//Success or failure loop for primitive backtrack search of augmented atom complex
while(1) {

	while(intMatchDepth < intBetaMatchCnt){
		while(intColCnt < intAlphaMatchCnt) {
			//If not used, select first available position.
			if(!intColsUsed[intColCnt]) {
				if((intMaxBipMatch[intMatchDepth][intColCnt] & (1<<intMatchDepth))) {

					intMaxBipMatch[intMatchDepth][intColCnt] |= (1<<(intMatchDepth+1));
                    intColsUsed[intColCnt]=1;
                    intColatDUsed[intMatchDepth]=intColCnt;
                    break;
                }
            }
            intColCnt++;
        }
        //Check to see if search is a complete success and increment to next depth.
        if(intColCnt < intAlphaMatchCnt) {
			//Search was a success;therefore, exit loop and goto next match listing element
			if(intMatchDepth == (intBetaMatchCnt-1)) return 1;
            
            intColCnt=0;
            intMatchDepth++;

            //Copy to next depth.
			for(i=intMatchDepth;i<intBetaMatchCnt;i++) {
				for(j=0;j<intAlphaMatchCnt;j++) {
					if(intMaxBipMatch[i][j]) intMaxBipMatch[i][j] |= ((intMaxBipMatch[i][j] & (1<<(intMatchDepth-1)))<<1);
                }
            }
            
		}
        //Check to see if search is a failure for the match listing element under consideration (not 1:1 correspondence) and backtrack a depth.
        else {
			//Bipartite matching failed;therefore, delete appropriate match listing element
			if(intMatchDepth==0) return 0;

			//Backtrack a depth.
            intMatchDepth--;
            intColCnt=intColatDUsed[intMatchDepth]+1;
            intColsUsed[intColCnt-1]=0;
            intColatDUsed[intMatchDepth]=-1;
        }
    }
}

}

//This is a completely dynamic version of the refinement procedure used in the Ullman subisomorph alogorithm.  
//Note in this version, the routine does not default to a maximum bipartite matching of the augmented atom complex 
//used in the refinement procedure.  It defaults to a simple scan for possible neighbor matches to each match listing,
//and if an ambiguity in the matching scan results, the routine uses a primitive backtrack method to determine if a maximum 
//bipartite matching results in a match between the augmented atoms complexes.
unsigned _int8
Dyn_Ull_Refine(MatchStructure *ptrMS,Atom *ptrAlphaGraph,Atom *ptrBetaGraph,unsigned _int32 intNumAlphaAtoms,unsigned _int32 intNumBetaAtoms,unsigned _int32 intBeta_Depth) {

unsigned _int32 intLocLastChange=0;     //Variable denoting location of last matrix change
unsigned _int32 intBeta_Row=0;       //Beta variable counter used in scanning ptrMatchList
unsigned _int32 intAlpha_Place=0;          //Location of alpha atom from match list in alpha structure
MatchElement *ptrTempMatch;				//Pointer to current match element in a match listing chain at a specified depth for a given beta atom
MatchElement *ptrTempCompare;			//Pointer to neighbor match elements of a match listing chain at a specified depth for a given beta atom
MatchElement *ptrOld_Match;				//Pointer to match element previous to ptrTempMatch pointer

//Augmented atom matrix used in maximum bipartite matching when matching ambiguitites arise 
unsigned _int8 intMaxBipMatch[6][6];

//These variables are used in the "direct screening procedure
unsigned _int8 intPossibleMatches;			//Bitfield integer used to detect possible matching redundancies (ambiguities)
unsigned _int8 intAlphaMatchCnt;			//Alpha variable counter for number of possible matches for match list
unsigned _int8 intBetaMatchCnt;				//Beta variable counter for number of possible matches for match list
unsigned _int8 intMB_Flag=0;				//Flag to detect if maximum bipartite matching is necessary
unsigned _int8 intSimpleScanFlag;			//Flag to determine if simple scan (no bipartite matching) failed
unsigned _int8 intComplete_Loop=0;			//Flag to determine whether a complete loop with no deletion or "racking" of a match element has occurred.

//These variables are used in the maximum bipartite matching when ambiguities are discovered
unsigned _int8 i;                       //Loop counter
unsigned _int8 j;                       //Loop counter
AtomBond *ptrTempAlpha;					//Temporary pointer to an alpha AtomBond
AtomBond *ptrTempBeta;					//Temporary pointer to a beta AtomBond

//Overall loop to determine if a complete loop through ptrMatchList[~][~] has been accomplished since last element change

do {

     //Determine whether match location lies in the "used" region or in the "copied to next level region"
     //(Use ptrMS->ptrCol_Loc[~] array because match lies in "used"
     if(intBeta_Row<intBeta_Depth) ptrTempMatch=ptrMS->ptrCol_Loc[intBeta_Row];
     //Use "un-investigated" region.
     else ptrTempMatch=ptrMS->ptrMS_Spine[intBeta_Row];
          
     ptrOld_Match=ptrTempMatch;
     //Refinement loop for current match listing
     do { 
          
          //Obtain current location in alpha graph of the match node under consideration
          if(ptrTempMatch) intAlpha_Place=ptrTempMatch->intMatch;
          //Catastrophic failure.  Do not procede with refinement.
		  else return 0;
		            
          //Initialize intMaxBipMatch[~][~] to zero for this matching if necessary
          for(i=0;i <6;i++) {
               for(j=0;j<6;j++) {
                    intMaxBipMatch[i][j]=0;
               }
          }
          
          //Initialize variables prior to beta loop
          intBetaMatchCnt=0;
          intPossibleMatches=0;
          
          ptrTempBeta=ptrBetaGraph[intBeta_Row].ptrNextBond;
          
          //LOOP THROUGH ATTACMENTS TO THE BETA ELEMENT IN MATCH LIST
          while(ptrTempBeta) {

               intAlphaMatchCnt=0;
               intSimpleScanFlag=0;
               ptrTempAlpha=ptrAlphaGraph[intAlpha_Place].ptrNextBond;

               //LOOP THROUGH ATTACHMENTS TO THE ALPHA ELEMENT IN MATCH LIST
               while(ptrTempAlpha) {

                    //If bonds between the alpha and beta elements correspond or the beta bond is a wild card bond
                    if((ptrTempBeta->Bond_ID.Bond_Type == ptrTempAlpha->Bond_ID.Bond_Type && !ptrTempBeta->Bond_ID.Not_Bond) || (ptrTempBeta->Bond_ID.Bond_Type != ptrTempAlpha->Bond_ID.Bond_Type && ptrTempBeta->Bond_ID.Not_Bond) || ptrTempBeta->Bond_ID.Bond_Type == 15){
                         
                         //ptrTempBeta->intAttachedAtom represents the "used" region in the match listing.
                         if((ptrTempBeta->intAttachedAtom)<intBeta_Depth) {
                              ptrTempCompare=ptrMS->ptrCol_Loc[ptrTempBeta->intAttachedAtom];
                         }
                         //ptrTempBeta->intAttachedAtom represents the "un-investigated" region in the match listing.
                         else {
                              ptrTempCompare=ptrMS->ptrMS_Spine[ptrTempBeta->intAttachedAtom];
                         }
                         
                         //LOOP THROUGH CHAIN OF ELEMENTS FOR THE GIVEN NEIGHBOR OF THE BETA ATOM SPECIFIED BY intBeta_Row
                         while(ptrTempCompare) {
                         
                              //If the element in the match list corresponds to one of the possible alpha elements in ptrAlphaGraph
                              if(ptrTempCompare->intMatch == ptrTempAlpha->intAttachedAtom){
                                   
                                   intSimpleScanFlag=1;
                                   //Check to see if bit location in intPossibleMatches is "virgin"
                                   if(!(intPossibleMatches & (1<<intAlphaMatchCnt))) {
                                        //Fill bit location with "1"
                                        intPossibleMatches |= (1<<intAlphaMatchCnt);
                                   }
                                   //Set flag to enter maximum bipartite matching algorithm
                                   else {intMB_Flag=1;}

                                   //Fill location in matrix intMaxBipMatch[~][~] for possible use in maximum bipartite matching
                                   intMaxBipMatch[intBetaMatchCnt][intAlphaMatchCnt]=1;
                                   break;
                              }
                         
                              //ptrTempCompare is in the "used" region.
                              if((ptrTempBeta->intAttachedAtom)<intBeta_Depth) {
                                   ptrTempCompare=0;
                              }
                              //ptrTempCompare is in the "un-investigated" region.
                              else {
                                   ptrTempCompare=ptrTempCompare->ptrNextMatch;
                              }
                         }
                    }
                    intAlphaMatchCnt++;
                    ptrTempAlpha=ptrTempAlpha->ptrNextBond;

               } //while(ptrTempAlpha)

			   //Check to see if simple scan resulted in a non-match (failure) for these augmented atom complexes.
               //If so, delete appropriate match listing element.
               if(!intSimpleScanFlag) {
                    
                    intLocLastChange = intBeta_Row;

                    //If current match element is not in the "used" zone and has an additional attachment.
                    if(intBeta_Row >= intBeta_Depth && ptrMS->ptrMS_Spine[intBeta_Row]->ptrNextMatch) {                   
                    
						//If not the initial refinement, then use the storage rack for refined elements;
						if(intBeta_Depth) {
							if(!ptrMS->Rack_MN(intBeta_Row,intBeta_Depth,ptrOld_Match,ptrTempMatch)) return 0;
						}
						//Else if initial refinement, then delete match nodes directly from the match structure as they cannot be valid nodes.
						else {
							if(!ptrTempMatch->Delete_MN(ptrOld_Match));
						}
						ptrTempMatch=ptrOld_Match;
						break;
                    }
                    //If current match element is in the "used" zone or is the only match element remaining in chain.
                    else return 0; //Refinement resulted in a catastrophic failure(no possible match for specified beta atom).
               }
			   intBetaMatchCnt++;
               ptrTempBeta=ptrTempBeta->ptrNextBond;
          } //while(ptrTempBeta)
          
		  //Simple scan resulted in an ambiguity; therefore, enter into maximum bipartite matching
          if(intMB_Flag && intSimpleScanFlag) {
			  //Perform perfect,maximum bipartite matching of ambiguous, augmented atom complex.
			  if(!DFS_Max_Bip(intAlphaMatchCnt,intBetaMatchCnt,intMaxBipMatch)) {
					intLocLastChange=intBeta_Row;
			
					//If current match element is not in the "used" zone and has an additional attachment.
                    if(intBeta_Row >= intBeta_Depth && ptrMS->ptrMS_Spine[intBeta_Row]->ptrNextMatch) {                   
                    
						//If not the initial refinement, then use the storage rack for refined elements;
						if(intBeta_Depth) {
							if(!ptrMS->Rack_MN(intBeta_Row,intBeta_Depth,ptrOld_Match,ptrTempMatch)) return 0;
						}
						//Else if initial refinement, then delete match nodes directly from the match structure as they cannot be valid nodes.
						else {
							if(!ptrTempMatch->Delete_MN(ptrOld_Match));
						}
						ptrTempMatch=ptrOld_Match;
                    }
                    //If current match element is in the "used" zone or is the only match element remaining in chain.
                    else return 0; //Refinement resulted in a catastrophic failure(no possible match for specified beta atom).			
              }
		  }	//if(intMB_Flag && intSimpleScanFlag) 
	
		 intMB_Flag=0;
		 ptrOld_Match=ptrTempMatch;
     
		 //If the current match node is in the "used" zone, discontinue comparing.
		 if(intBeta_Row<intBeta_Depth) ptrTempMatch=0;
 		 //Else continue to next match element node for refinement.
 		 else if(ptrTempMatch->intMatch == intAlpha_Place) ptrTempMatch=ptrTempMatch->ptrNextMatch;

     }while(ptrTempMatch);
     
     //Variable Counter logic to scan through the match list
     if(intBeta_Row < (intNumBetaAtoms-1)) {
          intBeta_Row++;
     }
     else {
          intBeta_Row=0;
     }
	 
	 //Completed a loop without a match element deletion
	 if(intComplete_Loop) break;
	 //Signal a complete loop through the match structure.
	 else if(intLocLastChange == intBeta_Row) intComplete_Loop=1;

}while(1);

//*************************Printout for debugging purposes
//MatchElement *ptrTemp;
//cout<<"post-refinement"<<endl;
//for(i=0;i<intNumBetaAtoms;i++) {
//   ptrTemp=ptrMS->ptrMS_Spine[i];
//   if(ptrTemp) {
//	   cout<<(ptrTemp->intMatch)<<"->";
                       
//       while(ptrTemp->ptrNextMatch) {
//	       ptrTemp=ptrTemp->ptrNextMatch;
//           cout<<(ptrTemp->intMatch)<<"->"; 
//       }
//   }
//   cout<<endl;
//}
//cout<<endl;
//******************************

//Refinement did not result in a catostrophic failure (i.e., no possible matches for any given beta atom).
return 1;

}

//This routine is a dynamic version of the initial matching used in the Ullman (1976) subgraph isomorphism algorithm.
//It can perform a truncating backtrack or a combinatorial enumerative search where all possibly overlapping subfragments are detected.  If the truncating option is selected
//then the detection locations are stored as integer locations in the alpha structure.  If the enumerative option is selected, then the detections are located as bits within an array
//representing the alpha structure.
unsigned _int8
Ull_BackTrack(MatchStructure *ptrMS,Atom *ptrAlphaGraph,Atom *ptrBetaGraph,unsigned _int32 intNumAlphaAtoms,unsigned _int32 intNumBetaAtoms,unsigned _int8 intSearch_Type) {

MatchElement *ptrTemp;						//Debug printing pointer
unsigned _int32 i;                          //Junk counter variable
unsigned _int32 intBeta_Row=0;              //Current row in match linked list representing which beta atom is being investigated
unsigned _int32 intBeta_Depth=0;			//Current depth of search for match listing
const unsigned _int8 constBitSize=32;		//Constant used to declare size of bit field integer used in storing detections during enumerating search
unsigned _int32 intTempArrayLoc;			//Temporarily stores integer location in alpha detection bit field for "bit setting"
unsigned _int8 intTempBitLoc;				//Temporarily stores bit location in integer in alpha detectionb bit field for "bit setting"

unsigned _int8 *ptrAlphaUsed=new unsigned _int8[intNumAlphaAtoms];    //Array storing which columns have been used in depth search
SubFragLoc *ptrTempAlphaLoc;				//Temporary subfragment detection array pointer
SubFragLoc *ptrOldAlphaLoc;                 //Temporary subfragment detection array pointer
SubFragLoc *ptrJunkAlphaLoc;				//Temporary subfragment detection array pointer

//Set initial internal ptrAlphaLoc pointer
ptrTempAlphaLoc=ptrMS->ptrAlpha_Loc;

//This initializes the ptrAlphaUsed[] array to zero
for(i=0;i<intNumAlphaAtoms;i++) {
     ptrAlphaUsed[i]=0;
}

//Call the refinement procedure to "cull-out" non-matches.  If the refinement procedure returns
//a zero, the refinement noted a failure for the search at the specified depth (i.e., no alpha
//atoms matched a beta atom at the given depth in the search).  If the refinement procedure 
//returned a 1, then the search has not resulted in a failure.
if(!Dyn_Ull_Refine(ptrMS,ptrAlphaGraph,ptrBetaGraph,intNumAlphaAtoms,intNumBetaAtoms,intBeta_Depth)) {
     delete[] ptrAlphaUsed;
     return 0;
}
else {

     //THIS BLOCK LOOP PERFORMS THE BACK TRACKING PROCEDURE WITH INTERMITTENT REFINEMENT
     while(1) {

          //Check if there is a match location in the ptrCol_Loc[~] array already. If so, replace with the next one.
          if(ptrMS->ptrCol_Loc[intBeta_Depth]) {
               ptrAlphaUsed[ptrMS->ptrCol_Loc[intBeta_Depth]->intMatch]=0;
               ptrMS->ptrCol_Loc[intBeta_Depth]=ptrMS->ptrCol_Loc[intBeta_Depth]->ptrNextMatch;
          }
          //If not place in ptrMS->ptrCol_Loc[~]
          else ptrMS->ptrCol_Loc[intBeta_Depth]=ptrMS->ptrMS_Spine[intBeta_Depth];
          
          //Make sure that location has not already been provisationally matched
          while(ptrMS->ptrCol_Loc[intBeta_Depth] && (ptrAlphaUsed[ptrMS->ptrCol_Loc[intBeta_Depth]->intMatch])) {
               ptrMS->ptrCol_Loc[intBeta_Depth]=ptrMS->ptrCol_Loc[intBeta_Depth]->ptrNextMatch;
          }
     
          //If ptrMS->ptrCol_Loc[~] is null because no element at given depth is available, check to see if it is a instance failure or total failure.
          if(!ptrMS->ptrCol_Loc[intBeta_Depth]) {
               //Instance failure. //Backtrack a depth and re-append the "stored" refined match element nodes.
			   if(intBeta_Depth) {
				   //Update ptrAlphaUsed[~].  Remove previous provisational match.
				   intBeta_Depth--;
				   ptrAlphaUsed[ptrMS->ptrCol_Loc[intBeta_Depth]->intMatch]=0;
                   ptrMS->UnRack_MN(intBeta_Depth,intNumBetaAtoms);
               }
               //Total failure
               else {
                     delete[] ptrAlphaUsed;
                    //Return either a failure or a success depending upon whether a previous subfragment was detected.
                    if(ptrTempAlphaLoc == ptrMS->ptrAlpha_Loc) return 0; //Return a complete search failure
                    else {
                         delete ptrOldAlphaLoc->ptrNextSubFragLoc;
                         ptrOldAlphaLoc->ptrNextSubFragLoc=0;
                         return 1;      //Return a search success because at least one subfragment was detected.
                    }
               }
          }
     
          //INCREASE A DEPTH
          else {

               ptrAlphaUsed[ptrMS->ptrCol_Loc[intBeta_Depth]->intMatch]=1;

               //If depth has reached the maximum, then return a successful detection
			   if(intBeta_Depth == (intNumBetaAtoms-1)) {
               
                    //If backtrack search type is set to non-enumerating option
                    if(intSearch_Type<2) {

                         //Fill the ptrTempAlphaLoc->ptrQuery_Loc[~] array with the integer array locations in the alpha graph
                         for(i=0;i<intNumBetaAtoms;i++){
                              ptrTempAlphaLoc->ptrQuery_Loc[i]=ptrMS->ptrCol_Loc[i]->intMatch;
                         }
                         
                         //Return a search success because a subfragment was detected.
                         delete[] ptrAlphaUsed;
                         return 1;
                    }
                    //Else if backtrack search type is set to enumerating option
                    else {
                         
                         //Fill the ptrTempAlphaLoc->ptrQuery_Loc[~] array with bit locations in the alpha bit field (an integer array where each bit represents in alpha atom).
                         for(i=0;i<intNumBetaAtoms;i++) {
                              intTempArrayLoc=(ptrMS->ptrCol_Loc[i]->intMatch) / constBitSize;
                              intTempBitLoc=(ptrMS->ptrCol_Loc[i]->intMatch) % constBitSize;

                              //If beta atom does not correspond with a "decharacterized" alpha atom
                              if(!ptrBetaGraph[i].Search_ID.Dechar_Atom) ptrTempAlphaLoc->ptrQuery_Loc[intTempArrayLoc] |= (1<<intTempBitLoc);
                         }
                         //Check for redundant occurrence of subfragment.  If found, then set flag for this listing to be erased (set to zero)
                         ptrJunkAlphaLoc=ptrMS->ptrAlpha_Loc;
                         while(ptrJunkAlphaLoc != ptrTempAlphaLoc) {

                              for(i=0;i<ptrMS->ptrAlpha_Loc->intArray_Length;i++) {
                                   if(ptrJunkAlphaLoc->ptrQuery_Loc[i]^ptrTempAlphaLoc->ptrQuery_Loc[i]) break;
                              }
                              //If the current subfragment listing is a duplicate, then reset the detection array equals zero.
                              if(i==ptrMS->ptrAlpha_Loc->intArray_Length){
                                   for(i=0;i<ptrMS->ptrAlpha_Loc->intArray_Length;i++) {
                                        ptrTempAlphaLoc->ptrQuery_Loc[i]=0;
                                   }
                                   break;
                              }
                              ptrJunkAlphaLoc=ptrJunkAlphaLoc->ptrNextSubFragLoc;
                         }

                         //Add an increment to the detected atom array, if detection is not a duplicate.
                         if(ptrJunkAlphaLoc==ptrTempAlphaLoc) {
                              ptrOldAlphaLoc=ptrTempAlphaLoc;
                              ptrTempAlphaLoc->ptrNextSubFragLoc=new SubFragLoc(ptrMS->ptrAlpha_Loc->intArray_Length);
                              ptrTempAlphaLoc=ptrTempAlphaLoc->ptrNextSubFragLoc;
                         }
                    }
               }
               else {
                    //Increase depth
                    intBeta_Depth++;
                                   
//*************************Printout for debugging purposes
//cout<<"matching after depth increment"<<endl;
//for(i=intBeta_Depth;i<intNumBetaAtoms;i++) {
//		ptrTemp=ptrMS->ptrMS_Spine[i];
//      if(ptrTemp) {
//             cout<<ptrTemp->intMatch<<"->";
//             while(ptrTemp->ptrNextMatch) {
//                  ptrTemp=ptrTemp->ptrNextMatch;
//                  cout<<ptrTemp->intMatch<<"->"; 
//             }
//      }
//	  cout<<endl;
//}
//cout<<endl;
//******************************

                    //Refine match listing at current depth to determine whether to backtrack at depth.
                    if(!Dyn_Ull_Refine(ptrMS,ptrAlphaGraph,ptrBetaGraph,intNumAlphaAtoms,intNumBetaAtoms,intBeta_Depth)) {
						//Backtrack a depth.
						//Update ptrAlphaUsed[~].  Remove previous provisational match.
						intBeta_Depth--;
						ptrAlphaUsed[ptrMS->ptrCol_Loc[intBeta_Depth]->intMatch]=0;
						ptrMS->UnRack_MN(intBeta_Depth,intNumBetaAtoms);					
					}
               }
          }
     }
}
}

//Routine to check if Ring Feature operator bits corresponding to a constrained subfragment correspond between an alpha and a beta structure.
unsigned _int8
Check_Ring_Features(SearchID Alpha_SID,SearchID Beta_SID) {

//For compound atom nodes of same atomic number (i.e., [C,c] )
if(Beta_SID.Not_Ring && Beta_SID.In_Ring) return 1;

//Check if ring features are specified (i.e., <<R>>,<<!R>>,<<Rr>>,c,etc.)
else if(Beta_SID.Not_Ring || Beta_SID.In_Ring) {

     //Processed alpha atom is not a member of a ring.
     if(Alpha_SID.Not_Ring) {
          //Check to see if beta atom is not in a ring system.
          if(Beta_SID.Not_Ring && Beta_SID.Ring_Type==3 && Beta_SID.Ring_Size==0 && !Beta_SID.Less_Than && !Beta_SID.Greater_Than) {
               return 1;
          }
          else return 0;
     }
     //Processed alpha atom is a member of a ring.
     else if(Alpha_SID.In_Ring) {
          //Check to see if beta atom is not in a ring system.
          if(Beta_SID.Not_Ring && Beta_SID.Ring_Type==3 && Beta_SID.Ring_Size==0 && !Beta_SID.Less_Than && !Beta_SID.Greater_Than) {
               return 1;
          }
          //Beta atom is in a ring or constrained not to be in a specified ring size and/or type.
          else {
               //Check if ring type properties of alpha and beta structure match.
               if(Beta_SID.Ring_Type & Alpha_SID.Ring_Type) {
                    //If less than operator was specified.
                    if(Beta_SID.Less_Than) {
                         //Compare.  Is # of alpha atoms less than # of beta atoms?
                         if(Alpha_SID.Ring_Size < Beta_SID.Ring_Size) {
                              //If not bit is specified, 
                              if(Beta_SID.Not_Ring) return 0;
                              else return 1;
                         }
                         else {
                              if(Beta_SID.Not_Ring) return 1;
                              else return 0;
                         }
                    }
                    //If greater than operator was specified.
                    else if(Beta_SID.Greater_Than) {
                         //Compare.  Is # of alpha atoms greater than # of beta atoms?
                         if(Alpha_SID.Ring_Size > Beta_SID.Ring_Size) {
                              //If not bit is specified, 
                              if(Beta_SID.Not_Ring) return 0;
                              else return 1;
                         }
                         else {
                              if(Beta_SID.Not_Ring) return 1;
                              else return 0;
                         }
                    }     
                    //Else if an exact # of ring elements were specified
                    else if(Beta_SID.Ring_Size) {
                         //If number of atoms match
                         if(Alpha_SID.Ring_Size==Beta_SID.Ring_Size) {
                              if(Beta_SID.Not_Ring) return 0;
                              else return 1;
                         }
                         //If number of atoms do not match
                         else {
                              if(Beta_SID.Not_Ring) return 1;
                              else return 0;
                         }
                    }
                    //Else no number of ring elements were specified 
                    else {
                         //If not bit is set.  
                         if(Beta_SID.Not_Ring) return 0;
                         //Must be in a ring. 
                         else return 1;
                    }     
               }
               //Ring types do not match.
               else return 0;
          }
     }
     //Alpha atom has not been processed, and beta atom has been declared in a ring implicitly by chemical string syntax.
     //(i.e., Alpha is -C-, and beta is -c-.)
     else return 0;
}
//Else ring specification has not been declared explicitly and may or may not be present implicitly (i.e., C,O,etc.)
else {
     if(!Alpha_SID.Not_Ring && !Alpha_SID.In_Ring) return 1;
     else return 0;
}

}     

//This routine creates an initial match listing structure based on a diagonal matrix of pointers that is used by the Ullman subgraphisomorphism algorithm.
//It is used in the DynamicUllman subgraphisomorphism routine.
unsigned _int8
Ullman_Subgraph(HashStructure *ptrHash_Table,MatchStructure *&ptrMS,Atom *ptrAlphaGraph,Atom *ptrBetaGraph,unsigned _int32 intNumAlphaAtoms,unsigned _int32 intNumBetaAtoms,unsigned _int8 intSearch_Type) {

unsigned _int32 i;                           //Junk counter variable
unsigned _int32 intBeta_Row;				 //Current row in match linked list representing which beta atom is being investigated
unsigned _int8 intMatchDetect=0;			 //Flag to determine if no possible match was found for a particular beta atom
unsigned _int32 intHashPlace;				 //Location in hash header of possible alpha matches for a particular beta atom
unsigned _int32 intStop;					 //Loop boundary variable
unsigned _int32 intLP;						 //Location in alpha structure during comparison
MatchElement *ptrTemp3;						 //Temporary pointer used in adding a match element for an initial match

//Check whether a subfragment has been properly loaded.
if(!(intNumBetaAtoms && intNumAlphaAtoms)) return 0;

//Create match structure entity.
if(!ptrMS) ptrMS=new MatchStructure(intNumBetaAtoms,intNumAlphaAtoms,intSearch_Type);

//THIS BLOCK ROUTINE FILLS THE MATCH LINKED LIST AT DEPTH ZERO PRIOR TO FIRST REFINEMENT
i=0;

//Note: The manner in which this block is formed is code redundant, but it was constructed
//in this manner to improve readability, ease possible future amendments to the current bitfield
//and reduce the number of unnecessary bit comparison operations.
for(intBeta_Row=0;intBeta_Row<intNumBetaAtoms;intBeta_Row++) {

	 ptrTemp3=ptrMS->ptrMS_Spine[intBeta_Row];
     
     //If the beta atom is an unspecified atom
     if(ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Atom) {
     
          //If the atom is decharacterized (therefore, doesn't matter if it has an unspecifed
          //connection or not), use this if block.
          if(ptrBetaGraph[intBeta_Row].Search_ID.Dechar_Atom) {

               //If there is an unspecified atom connection
               if(ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Neighbor){

                    //Loop to step through ptrAlphaGraph because an unspecified character precludes the
                    //use of the hashing function.
                    while(i<intNumAlphaAtoms) {
                         						
						 //Check ring features
                         if(Check_Ring_Features(ptrAlphaGraph[intLP].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
                              ptrMS->New_MN(ptrTemp3,intBeta_Row,i);
                              intMatchDetect=1;   //Flag a match for this beta atom
                         }
                         i++;
                    }
               }
               //If there is no unspecified connection
               else {
                         
                    //Loop to step through ptrAlphaGraph because an unspecified character precludes the
                    //use of the hashing function.
                    while(i<intNumAlphaAtoms) {
						 
                         //Check ring features
                         if(Check_Ring_Features(ptrAlphaGraph[i].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
                              ptrMS->New_MN(ptrTemp3,intBeta_Row,i);
                              intMatchDetect=1;   //Flag a match for this beta atom
                         }
                         i++;
                    }
               }

          }
          //If atom is not a decharacterized atom
          else { 
               //If atom has an unspecified atom connection
               if(ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Neighbor) {
                         
                    //Loop to step through ptrAlphaGraph because an unspecified character precludes the
                    //use of the hashing function.
                    while(i<intNumAlphaAtoms) {

                         //If ptrAlphaGraph[~] atom does not have its "already used" flag set
                         if(!ptrAlphaGraph[i].Search_ID.Used_Atom) {
                         
                              //If unspecified bond connection bit is not set
                              if(!ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Bond) {

                                   //If the # of non-hydrogen connections, and bond type bits correspond
                                   if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[i].Atom_ID.Num_NH_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Bond_Types==ptrAlphaGraph[i].Atom_ID.Bond_Types) {
										//Check ring features
                                        if(Check_Ring_Features(ptrAlphaGraph[i].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
                                             ptrMS->New_MN(ptrTemp3,intBeta_Row,i);
                                             intMatchDetect=1;   //Flag a match for this beta atom
                                        }  
                                   }
                              }
                              //If unspecified bond connection flag is set
                              else {

                                   //Compare # non-hydrogen connections
                                   if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[i].Atom_ID.Num_NH_Cnt) {
                                        //Check ring features
                                        if(Check_Ring_Features(ptrAlphaGraph[i].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
                                             ptrMS->New_MN(ptrTemp3,intBeta_Row,i);
                                             intMatchDetect=1;   //Flag a match for this beta atom
                                        }   
                                   }
                              }
                         }
                         i++;
                    }
               }
               //If atom does not have an unspecified atom connection
               else {
                         
                    //Loop to step through ptrAlphaGraph because an unspecified character precludes the
                    //use of the hashing function.
                    while(i<intNumAlphaAtoms) {
                         //If ptrAlphaGraph[~] atom does not have its "already used" flag set
                         if(!ptrAlphaGraph[i].Search_ID.Used_Atom) {

                              //If unspecified bond connection flag is not set
                              if(!ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Bond) {

                                   //Compare # non-hydrogen connections, atom charge, bond types, atom connection types.
								  if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[i].Atom_ID.Num_NH_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Bond_Types==ptrAlphaGraph[i].Atom_ID.Bond_Types && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Charge==ptrAlphaGraph[i].Atom_ID.Atom_Charge && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Cnt==ptrAlphaGraph[i].Atom_ID.Atom_Cnt) {
                                        
                                        //Check ring features
                                        if(Check_Ring_Features(ptrAlphaGraph[i].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
                                             ptrMS->New_MN(ptrTemp3,intBeta_Row,i);
                                             intMatchDetect=1;   //Flag a match for this beta atom
                                        }   
                                   }
                              }    
                              //If unspecified bond connection flag is set
                              else {
                                   //Compare # non-hydgrogen, atom charge, and atom connection type bits
                                   if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[i].Atom_ID.Num_NH_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Charge==ptrAlphaGraph[i].Atom_ID.Atom_Charge && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Charge==ptrAlphaGraph[i].Atom_ID.Atom_Charge && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Cnt==ptrAlphaGraph[i].Atom_ID.Atom_Cnt){
                                
										//Check ring features
                                        if(Check_Ring_Features(ptrAlphaGraph[i].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
                                             ptrMS->New_MN(ptrTemp3,intBeta_Row,i);
                                             intMatchDetect=1;   //Flag a match for this beta atom
                                        }   
                                   }
                              }
                         }
                         i++;
                    }
               }
          }
          
     }
          
     //If the beta atom is not an unspecifed atom.
     else {
          //If atom is decharacterized
          if(ptrBetaGraph[intBeta_Row].Search_ID.Dechar_Atom) {

			   //Determine alpha loop length (whether hashing is used or not)
			   if(ptrHash_Table && ptrHash_Table->intHash_Flag) {
				  //Temporarily store hash result of atom type.
				  intHashPlace=ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type % ptrHash_Table->intHash_PN;
				  intStop=ptrHash_Table->ptrHash_Bucket[intHashPlace].intChainLength;
			   }
			   else intStop=intNumAlphaAtoms;

               //If atom has an unspecified atom connection
               if(ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Neighbor) {

                    while(i<intStop) {
          
                         //Get location in alpha graph
                         if(ptrHash_Table && ptrHash_Table->intHash_Flag) intLP=*(ptrHash_Table->ptrHash_Bucket[intHashPlace].ptrSepChain+i);
                         else intLP=i;

                         //Check ring features
                         if(Check_Ring_Features(ptrAlphaGraph[intLP].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
							//If atom types correspond
							if(ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type==ptrAlphaGraph[intLP].Atom_ID.Atom_Type) {
								//If not atom bit is not set.
                                if(!ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
									ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
                                    intMatchDetect=1;   //Flag a match for this beta atom
                                }
                            }   
							//Atom types do not ocrrespond.
							else if(ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
								ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
                                intMatchDetect=1;   //Flag a match for this beta atom
							}
                         }
						 i++;
                    }
               }
               //If atom has no unspecified atom connection
               else {
                    
                    while(i<intStop) {
                    
                         //Get location in alpha graph
                         if(ptrHash_Table && ptrHash_Table->intHash_Flag) intLP=*(ptrHash_Table->ptrHash_Bucket[intHashPlace].ptrSepChain+i);
                         else intLP=i;
                         
						 //Check ring features
                         if(Check_Ring_Features(ptrAlphaGraph[intLP].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {	
							//If atom types correspond
							if(ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type==ptrAlphaGraph[intLP].Atom_ID.Atom_Type) {
								//If not atom bit is not set.
                                if(!ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
									ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
                                    intMatchDetect=1;   //Flag a match for this beta atom
                                }
                            }   
							//Atom types do not correspond.
							else if(ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
								ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
                                intMatchDetect=1;   //Flag a match for this beta atom
							}
                         }
                         i++;
                    }
               }
          }
          //If atom is not a decharacterized atom
          else {
			   //Determine alpha loop length (whether hashing is used or not)
			   if(ptrHash_Table && ptrHash_Table->intHash_Flag) {
				  //Temporarily store hash result of atom type.
				  intHashPlace=(ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type | (ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt<<7)) % ptrHash_Table->intHash_PN;
				  intStop=ptrHash_Table->ptrHash_Bucket[intHashPlace].intChainLength;
			   }
			   else intStop=intNumAlphaAtoms;

			   //If atom has an unspecified atom connection
               if(ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Neighbor) {
                    
                    while(i<intStop) {
               
                         //Get location in alpha graph
                         if(ptrHash_Table && ptrHash_Table->intHash_Flag) intLP=*(ptrHash_Table->ptrHash_Bucket[intHashPlace].ptrSepChain+i);
                         else intLP=i;

                         //If ptrAlphaGraph[~] atom does not have its "already used" flag set
                         if(!ptrAlphaGraph[intLP].Search_ID.Used_Atom) {
							  
							 //If unspecified bond connection flag is not set
                              if(!ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Bond) {
								   
								   //Check ring features
                                   if(Check_Ring_Features(ptrAlphaGraph[intLP].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
								  
										//If atom types correspond.
									   if(ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type==ptrAlphaGraph[intLP].Atom_ID.Atom_Type){
										   //If # of non-hydrogen connections, atom charge, and bond type connection bits correspond
										   if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[intLP].Atom_ID.Num_NH_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Charge==ptrAlphaGraph[intLP].Atom_ID.Atom_Charge && ptrBetaGraph[intBeta_Row].Atom_ID.Bond_Types==ptrAlphaGraph[intLP].Atom_ID.Bond_Types){
												//If not atom bit is not set.
												if(!ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
													ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
													intMatchDetect=1;   //Flag a match for this beta atom
												}
										   }   
									   }
									   //Atom types do not correspond.
									   else if(ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
										   ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
										   intMatchDetect=1;   //Flag a match for this beta atom
									   }
								   }

                              }
                              //If unspecified bond connection flag is set
                              else {

								   //Check ring features
                                   if(Check_Ring_Features(ptrAlphaGraph[intLP].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {	

										//Atom types correspond.
										if(ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type==ptrAlphaGraph[intLP].Atom_ID.Atom_Type) {
											//If # non-hydrogen connections, atom charge correspond.
											if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[intLP].Atom_ID.Num_NH_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Charge==ptrAlphaGraph[intLP].Atom_ID.Atom_Charge) {
												//If not atom bit is not set.
												if(!ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
													ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
													intMatchDetect=1;   //Flag a match for this beta atom
												}
											}   
										}
										//Atom types do not correspond.
										else if(ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
											ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
											intMatchDetect=1;   //Flag a match for this beta atom
										}
								   }
                              }
                         }
                         i++;
                    }
                    
               }
               //If atom has no unspecified atom connection
               else {
                         
                    while(i<intStop) {
                    
                         //Get location in alpha graph
                         if(ptrHash_Table && ptrHash_Table->intHash_Flag) intLP=*(ptrHash_Table->ptrHash_Bucket[intHashPlace].ptrSepChain+i);
                         else intLP=i;

                         //If ptrAlphaGraph[~] atom does not have its "already used" flag set
                         if(!ptrAlphaGraph[intLP].Search_ID.Used_Atom) {
                              
                              //If unspecified bond connection flag is not set
                              if(!ptrBetaGraph[intBeta_Row].Search_ID.Unspec_Bond) {

									//Check ring features
                                    if(Check_Ring_Features(ptrAlphaGraph[intLP].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {	
										//Atom types correspond.
										if(ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type==ptrAlphaGraph[intLP].Atom_ID.Atom_Type) {
											//Compare # non-H's, atom charge, bond types, atom connection types.
											if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[intLP].Atom_ID.Num_NH_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Charge==ptrAlphaGraph[intLP].Atom_ID.Atom_Charge && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Cnt==ptrAlphaGraph[intLP].Atom_ID.Atom_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Bond_Types==ptrAlphaGraph[intLP].Atom_ID.Bond_Types){
												//If not atom bit is not set.
												if(!ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
													ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
													intMatchDetect=1;   //Flag a match for this beta atom
												}
											}   
										}
										//Atom types do not correspond.
										else if(ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
											ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
											intMatchDetect=1;   //Flag a match for this beta atom								
										}
									}

                              }
                              //If unspecified bond connection flag is set
                              else {

                                   //Check ring features
                                   if(Check_Ring_Features(ptrAlphaGraph[intLP].Search_ID,ptrBetaGraph[intBeta_Row].Search_ID)) {
										//Atom types correspond.
										if(ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Type==ptrAlphaGraph[intLP].Atom_ID.Atom_Type) {
											//Compare atom type, atom charge, # non-hydrogen attachments, and atom type connection bits
											if(ptrBetaGraph[intBeta_Row].Atom_ID.Num_NH_Cnt==ptrAlphaGraph[intLP].Atom_ID.Num_NH_Cnt && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Charge==ptrAlphaGraph[intLP].Atom_ID.Atom_Charge && ptrBetaGraph[intBeta_Row].Atom_ID.Atom_Cnt==ptrAlphaGraph[intLP].Atom_ID.Atom_Cnt){
												//If not atom bit is not set.
												if(!ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
													ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
													intMatchDetect=1;   //Flag a match for this beta atom
												}
											}   
										}
										//Atom types do not correspond.
										else if(ptrBetaGraph[intBeta_Row].Search_ID.Not_Atom) {
											ptrMS->New_MN(ptrTemp3,intBeta_Row,intLP);
											intMatchDetect=1;   //Flag a match for this beta atom
										}
									}
							  }
                         }
                         i++;
                    }
                                                       
               }
          }
     }

     //Check to see if no possible alpha matches were found for specified beta atom
     if(intMatchDetect) intMatchDetect=0;
     //No alpha matches were found for beta atom.  Exit search as failure.
     else return 0;
          
     //Reset i counter to zero for while loops
     i=0;

//Initial match loop closure
}    

//*************************Printout for debugging purposes
//MatchElement *ptrTemp;
//cout<<"initial match"<<endl;
//for(i=0;i<intNumBetaAtoms;i++) {
//   ptrTemp=ptrMS->ptrMS_Spine[i];
//   if(ptrTemp) {
//	   cout<<(ptrTemp->intMatch)<<"->";

//       while(ptrTemp->ptrNextMatch) {
//	       ptrTemp=ptrTemp->ptrNextMatch;
//           cout<<(ptrTemp->intMatch)<<"->"; 
//       }
//   }
//   cout<<endl;
//}
//cout<<endl;
//******************************

//If initial match did not result in a failure then call the Ull_BackTrack routine to perform the dynamic implementation of the Ullman algorithm.
if(Ull_BackTrack(ptrMS,ptrAlphaGraph,ptrBetaGraph,intNumAlphaAtoms,intNumBetaAtoms,intSearch_Type)) {
     //Search was a success.
     return 1;
}
//Search was a failure.
else return 0;

}

//This routine prints out to a specified file the results of a file query search.
inline void
MO_File_Output(ChemComboID *ptrAlphaIDNode,_int8 intSearchResult,ofstream &OutputAlphaFile,_int32 intSF_ID[100],_int32 intSF_Quant[100],_int32 intMF_ID[21],_int32 intMF_Quant[21]) {

	unsigned _int16 i;		//Junk counter.

	//Output the query structure string.
	OutputAlphaFile<<"\""<<ptrAlphaIDNode->chrChemSyntaxString<<"\""<<endl;
	
	//Output row title (Subfragment ID)
	OutputAlphaFile<<"Detection ID: "<<",";

	//Output search detection IDs if search was a success.
	if(intSearchResult==1) {
	
		//Output molecular feature ID
		if(intMF_ID[0]) {
			OutputAlphaFile<<"<,";
			for(i=0;i<21;i++) {
				if(intMF_ID[i] > 0) OutputAlphaFile<<intMF_ID[i]<<",";
				else break;
			} 
			OutputAlphaFile<<",>,";
		}
		//Output subfragment ID
		for(i=0;i<100;i++) {
			if(intSF_ID[i] > 0) OutputAlphaFile<<intSF_ID[i]<<",";
			else if(intSF_ID[i] <0) OutputAlphaFile<<"|,";
			else break;
		}
	}
	OutputAlphaFile<<endl;

	//Output row title (detection quantity).
	OutputAlphaFile<<"Detection Quant.: "<<","; 

	//Output search detection quantities if search was a success.
	if(intSearchResult==1) {
	
		//Output molecular feature quantity
		if(intMF_ID[0]) {
			OutputAlphaFile<<"<,";
			for(i=0;i<21;i++) {
				if(intMF_ID[i] > 0) OutputAlphaFile<<intMF_Quant[i]<<",";
				else break;
			} 
			OutputAlphaFile<<",>,";
		}
		//Output subfragment ID
		for(i=0;i<100;i++) {
			if(intSF_ID[i] > 0) OutputAlphaFile<<intSF_Quant[i]<<",";
			else if(intSF_ID[i] <0) OutputAlphaFile<<"|,";
			else break;
		}
	}
	OutputAlphaFile<<endl;
}

//This routine determines the proper return value depending upon the search results (intSearchResult) for the Molecular_Op() routine.
inline void
MO_Return_Values(_int8 intSearchType,unsigned _int8 intSuccessSearch,_int8 &intSearchResult) {

//If non-truncating search option is selected
if(intSearchType==0) {

	//Complete success (a subfragment was found during non-truncating search)
    if(intSuccessSearch) intSearchResult=1;
	//Complete failure.
	else intSearchResult=0;
}
//If truncating search option is selected.
else if(intSearchType==1) {
     
	if(intSuccessSearch) {
	//Partial success (truncating search resulted in a residual fragment(s)
		if(!intSearchResult) intSearchResult=2;
    }
    //Complete failure.
	else intSearchResult=0;
}
//If combinatorial truncating option is selected
else if(intSearchType==2) {
	if(intSuccessSearch) intSearchResult=1;
    else intSearchResult=0;
}

}

//This routine performs the subgraph isomorph operation.  It is designed to be called from a *.dll file
//or a main routine in an executable file.  This routine accepts a query SMILES structure, a file name
//of pseudo-SMILES subfragments, the type of search to be performed, and two arrays to store the ID of each detected 
//subfragment and its resective number of detections within the query structure.
void __declspec(dllexport) _stdcall 
MOSDAP(char *strQuery,_int8 intQueryType,char *strSubFragFile,char *strOutputQueryFile,_int8 intSearchType,_int8 &intSearchResult,_int32 intSF_ID[100],_int32 intSF_Quant[100],_int32 intMF_ID[21],_int32 intMF_Quant[21]) {	

MatchStructure *ptrMS=0;					 //Pointer to subgraphisomorph algorithm match structure.
QK_BF ptrTemp_Alpha_QK;						//Temporary atom quantity bit field used for "hypothetical" screening in non-truncating searches
unsigned _int32 intUsedAlphaAtoms;          //Counter used to determine if a subfragment has been completely "used" in a seq. trunc. search
char *ptrTmpAlphaString=0;                  //Char pointer used to read in alpha structures
unsigned _int32 j;                          //Junk micro-loop counter variable (one or less nested loops)
unsigned _int8 intNumScreenPockets=12;      //Number of bins(pockets) for bin sort & first-tier screening
unsigned _int8 intAlphaBinLoc;              //Counter for location in the 1st tier sort bin array
unsigned _int32 intSubFragDetect;           //Number of unique subfragment detections (excluding multiples of the same subfragment)
unsigned _int32 intMF_Detect;				//Number of unique molecular feature detections.
unsigned _int8 intNewSubFrag;				//Flag determining whether subfragment detection is unique or a multple detection
unsigned _int8 intSuccessSearch;			//Variable to detect whether a successful search has occurred
ComboLoc *ptrTempComboLoc=0;					//Temporary ComboLoc pointer
SubFragLoc *ptrTemp_SF_Loc=0;					//Temporary SubFragLoc pointer
HashStructure *ptrHash_Table=0;				//Hash table structure used in abbreviated initial atom matching declaration.

//Beta string list declarations
SubFragList *ptrBetaIDList;
ptrBetaIDList=new SubFragList;
ChemSeqID *ptrBetaIDNode;

//Alpha ID Node declarations
ChemComboID *ptrAlphaIDNode;
ptrAlphaIDNode=new ChemComboID;

//Declare query file
ifstream InputAlphaFile;
ofstream OutputAlphaFile;

//Read in subfragment file. If file is empty, return a search failure.
if(!(ptrBetaIDList->Read_SF_List(strSubFragFile,intQueryType))) return;

//If "file search" option is selected, then bin sort list into pockets
if(intQueryType) {
     ptrBetaIDList->ListBinSort(intNumScreenPockets);

	 //Open Query Alpha structure file
	 OutputAlphaFile.open(strOutputQueryFile);	

     //Open Query Alpha structure file
     InputAlphaFile.open(strQuery);
     ptrTmpAlphaString=new char[610];
     InputAlphaFile>>ptrTmpAlphaString;
     ptrAlphaIDNode->chrChemSyntaxString=new char[strlen(ptrTmpAlphaString)+1];
	 strcpy(ptrAlphaIDNode->chrChemSyntaxString,ptrTmpAlphaString);
}
//Is a single molecule query
else {
    ptrAlphaIDNode->chrChemSyntaxString=new char[strlen(strQuery)+1]; 
	strcpy(ptrAlphaIDNode->chrChemSyntaxString,strQuery);
}

//LOOP TO STEP THROUGH ALPHA STRING FILE
while(1) {

     //Search Return variables
	 intUsedAlphaAtoms=0;
	 intSuccessSearch=0;
     intSearchResult=0;
	 //Subfragment detection variables
     intSubFragDetect=0;
     intNewSubFrag=1;
	 //Molecular feature detection variables
	 intMF_Detect=0;

     //Perform atom quantity screen bit fill on alpha ID node
     ptrAlphaIDNode->SMILES_QK_Fill();

     if(intQueryType) {

          intAlphaBinLoc=0;

          //If "file search" option is selected, then fill the Occupancy Key for simple XOR bit operations
          ptrAlphaIDNode->QK_to_OK();
          
          //Increment through bin array until a location is found in the alpha occupation key, while the counter is less than the number of pockets, and the bin pointer is not null
          while((!(ptrAlphaIDNode->intOccupancyKey & (1<<intAlphaBinLoc)) && (intAlphaBinLoc<intNumScreenPockets)) || (ptrBetaIDList->ptrBinSortLoc[intAlphaBinLoc+1] == ptrBetaIDList->ptrBinSortLoc[intAlphaBinLoc])) {
               intAlphaBinLoc++;
          }

          //If an appropriate location was found.
          if(intAlphaBinLoc<intNumScreenPockets) ptrBetaIDNode=ptrBetaIDList->ptrBinSortLoc[intAlphaBinLoc];
          //If no appropriate bin for this alpha was found.
          else ptrBetaIDNode=0;
	 }
     //Use sequential pass through Beta list (single query structure option)
     else ptrBetaIDNode=ptrBetaIDList->ptrFirstIDNode;
     
     //This tallies the detected atoms as "set bits". If all bits are not set after all subfragment detections, then no need to proceed to costly exact cover algorithm.
     if(intSearchType==2) ptrAlphaIDNode->Initialize_EC_Check();

     //SCREENING & SUBSTRUCTURE SEARCH LOOP
     while(ptrBetaIDNode){

		  //Perform occupancy key screen if "file search" option is selected
          if(!(ptrAlphaIDNode->intOccupancyKey) || ((ptrAlphaIDNode->intOccupancyKey | ptrBetaIDNode->intOccupancyKey)==ptrAlphaIDNode->intOccupancyKey)) {

               //Check for presence of molecular feature and fill MF and appropriate QK bitfields if present.
			   if(!ptrBetaIDNode->Flags.QK_Fill) ptrBetaIDNode->Fill_MF();
			   
			   //Fill the pre-screen bitfield for subfragment structure.
			   if(!ptrBetaIDNode->Flags.QK_Fill) ptrBetaIDNode->SMILES_QK_Fill();
			   
			   //If average number of subfragment atoms has not been determined during file input, then set to size of current subfragment.
			   if(!ptrBetaIDList->intAvg_Num_Atoms) ptrBetaIDList->intAvg_Num_Atoms=ptrBetaIDNode->intNumberAtoms;	

               //Quantity screen the occupancy screened subfragment and master structures to pre-empt unneccessary substructure searching
               if(QK_Screen(ptrAlphaIDNode->intQuantKey,ptrBetaIDNode->intQuantKey)) {

                    //If alpha molecule has not already been processed, then process new molecule
                    if(!ptrAlphaIDNode->ptrMolecule) {
						 ptrAlphaIDNode->Parse_SMILES();
                         //Fill a hash header for the master structure to expedite substructure searching if it meets the iteration boundary condition.
                         if(0.07*ptrBetaIDList->intNumber_SF*ptrBetaIDList->intAvg_Num_Atoms>40) {
							 //Alpha hash table declarations
							 if(!ptrHash_Table) ptrHash_Table=new HashStructure;
							 ptrHash_Table->Hash_Graph(ptrAlphaIDNode->ptrMolecule->ptrAtom,ptrAlphaIDNode->intNumberAtoms);
						 }
                    }

                    //If beta string is a Molecular Feature string, then process beta string as a molecular feature instead of as a substructure.
                    if(ptrBetaIDNode->Flags.MF_Fill) {
                         
                         //If Bond Quantity specifiers are present, process as a bond feature.
                         if(ptrBetaIDNode->MF_ID.Bond_Feature) {
							intMF_Quant[intMF_Detect]=ptrAlphaIDNode->Retrieve_Bond_MF(ptrBetaIDNode);
							if(intMF_Quant[intMF_Detect]) {
								intMF_ID[intMF_Detect]=ptrBetaIDNode->intChemEntryID;
								intMF_Detect++;
								intSuccessSearch=1;
							}
                         }
                         //If Ring Quantity specifiers are present, process as a ring feature.
                         else if(ptrBetaIDNode->MF_ID.Ring_Feature) {

							intMF_Quant[intMF_Detect]=ptrAlphaIDNode->Retrieve_Ring_MF(ptrBetaIDNode);
							if(intMF_Quant[intMF_Detect]) {
								intMF_ID[intMF_Detect]=ptrBetaIDNode->intChemEntryID;
								intMF_Detect++;
								intSuccessSearch=1;
							}
                         }
                    }
                    
                    //If beta string has not already been processed as either a molecular subfragment or a molecular feature, then process new molecule
                    else {
                        if(!ptrBetaIDNode->ptrMolecule) ptrBetaIDNode->Parse_SMILES();

                        //Process Alpha structure if constraining molecular features associated with a beta subfragment are present, then proceed to substructure search using MF constraints.
						if(ptrBetaIDNode->Flags.RSC_Fill && !ptrAlphaIDNode->Flags.RSC_Fill) ptrAlphaIDNode->Calc_RSC();                    

                        //Loop to determine if initial substructure search for subfragment resulted in a success.  
                        if(Ullman_Subgraph(ptrHash_Table,ptrMS,(ptrAlphaIDNode)->ptrMolecule->ptrAtom,(ptrBetaIDNode)->ptrMolecule->ptrAtom,ptrAlphaIDNode->intNumberAtoms,ptrBetaIDNode->intNumberAtoms,intSearchType)) {

                              //If non-truncating search option is selected, then fill the temporary alpha quantity bit field for "temporary" screening
                              if(intSearchType==0) ptrAlphaIDNode->intQuantKey.Copy_QK(ptrTemp_Alpha_QK);

                              //Search loop for multiple subfragment occurrences
                              do {

                                   //If non-enumerating searches are selected, fill export arrays directly.
                                   if(intSearchType<2) {
                                        //Increment number of "unique" subfragment detections
                                        if(intNewSubFrag) intSubFragDetect++;
                                        intNewSubFrag=0;

                                        //Update the detected subfragment ID and increment the quantity
                                        intSF_ID[intSubFragDetect-1]=ptrBetaIDNode->intChemEntryID;
                                        intSF_Quant[intSubFragDetect-1]++;

                                        intSuccessSearch=1;
                                   }

                                   //NOTE: ptrAlphaLoc in search options 0 & 1 is recycled. The ptrQuery_Loc is not deleted or erased until the searching for the current alpha graph is complete. It is
								   //simply written over with each new detection. In search option 2, it is deleted after a search event and transfer of the ptrQuery_Loc[~] array to ptrQuery_Loc.

                                   //If the SEQUENTIAL TRUNCATING SEARCH type flag is set.
                                   if(intSearchType==1){

                                        //Truncate the Alpha Graph by setting the "already used" bit in the alpha molecule (for intSearchType=1 only)
                                        Truncate_Molecule(ptrMS,ptrAlphaIDNode,ptrBetaIDNode,intUsedAlphaAtoms);
                                        
                                        //Truncate the Atom Quantity screen bit for the alpha graph to reflect subfragment detections
                                        ptrAlphaIDNode->intQuantKey.Truncate_QK(ptrMS->ptrAlpha_Loc,ptrAlphaIDNode->ptrMolecule->ptrAtom,ptrBetaIDNode->ptrMolecule->ptrAtom,ptrBetaIDNode->intNumberAtoms);
          
                                        //Pre-screen the truncated alpha graph and beta graph structures to pre-empt unneccessary substructure searching.
                                        //If screen succeeds, then no need to do "already used" check.
                                        if(!QK_Screen(ptrAlphaIDNode->intQuantKey,ptrBetaIDNode->intQuantKey)) break;
                                                                           
                                        //Un-rack match elements from "storage rack" back onto the match structure spine.
										ptrMS->UnRack_All(ptrBetaIDNode->intNumberAtoms);
										
										//Delete detections from initial matching.
                                        if(!ptrMS->Truncate_Spine(ptrAlphaIDNode,ptrBetaIDNode)) break;

                                   }
                                   //If SEQUENTIAL NON-TRUNCATING SEARCH option is selected.
                                   else if(intSearchType==0) {
                                   
                                        //Truncate the Atom Quantity screen bit for the temporary alpha graph to reflect subfragment detections
                                        ptrTemp_Alpha_QK.Truncate_QK(ptrMS->ptrAlpha_Loc,ptrAlphaIDNode->ptrMolecule->ptrAtom,ptrBetaIDNode->ptrMolecule->ptrAtom,ptrBetaIDNode->intNumberAtoms); 
									   
                                        //Pre-screen the temporary truncated alpha graph and beta graph structures to pre-empt unneccessary substructure searching.
                                        //If pre-screen fails, then exit backtrack loop for current beta subfragment.
                                        if(!QK_Screen(ptrTemp_Alpha_QK,ptrBetaIDNode->intQuantKey)) break;
                                        
										//Un-rack match elements from "storage rack" back onto the match structure spine.
										ptrMS->UnRack_All(ptrBetaIDNode->intNumberAtoms);
										
										//Delete detections from initial matching.
										if(!ptrMS->Truncate_Spine(ptrAlphaIDNode,ptrBetaIDNode)) break;

                                   }
                                   //COMBINATORIAL (ENUMERATING) SEARCH option is selected.
                                   else if(intSearchType==2) {
                              
                                        //If current ComboLoc element is not the first in the list, then append a new one to the list.
                                        if(ptrAlphaIDNode->ptrFirstComboLoc) {
                                             ptrTempComboLoc->ptrNextComboLoc=new ComboLoc;
                                             ptrTempComboLoc=ptrTempComboLoc->ptrNextComboLoc; 
                                        }
                                        //Else if it is the first, then create the first one.
                                        else {
                                             ptrAlphaIDNode->ptrFirstComboLoc=new ComboLoc;
                                             ptrTempComboLoc=ptrAlphaIDNode->ptrFirstComboLoc;
                                        }
                                   
                                        ptrTemp_SF_Loc=ptrMS->ptrAlpha_Loc;
                                        while(ptrTemp_SF_Loc) {
                                   
                                             //Extract ptrQuery_Loc[~] from ptrTemp_SF_Loc and place as ptrQuery_Loc in ptrTempComboLoc
                                             ptrTempComboLoc->ptrQuery_Loc=ptrTemp_SF_Loc->ptrQuery_Loc;
                                             ptrTemp_SF_Loc->ptrQuery_Loc=0;
                                        
                                             //Bitwise OR the subfragment detection onto the exact cover check structure.
											 ptrAlphaIDNode->Fill_EC_Check(ptrTempComboLoc);
											                                    
                                             //Store address of detected beta ID in Combo element
                                             ptrTempComboLoc->ptrChemID=ptrBetaIDNode;
                                             //Create a new Combo element for multiple detections of subfragment.
                                             if(ptrTemp_SF_Loc->ptrNextSubFragLoc){
                                                  ptrTempComboLoc->ptrNextComboLoc=new ComboLoc;
                                                  ptrTempComboLoc=ptrTempComboLoc->ptrNextComboLoc;
                                             }
                                             ptrTemp_SF_Loc=ptrTemp_SF_Loc->ptrNextSubFragLoc;
                                        }
                                        break;
                                   }                                     
							  //Clear Alpha_Loc and ptrCol_Loc arrays in the match structure.
							  ptrMS->Initialize_Arrays(ptrBetaIDNode->intNumberAtoms);
							  
							  //Closure of multiple search loop
                              }while(Ull_BackTrack(ptrMS,ptrAlphaIDNode->ptrMolecule->ptrAtom,ptrBetaIDNode->ptrMolecule->ptrAtom,ptrAlphaIDNode->intNumberAtoms,ptrBetaIDNode->intNumberAtoms,intSearchType));
                         }    
                              
						if(ptrMS) {
							ptrMS->Clear_MS(ptrBetaIDNode->intNumberAtoms);
							delete ptrMS;
						}
						ptrMS=0;

                    }//End of MF / SF decision if/else if block                    
                    
                    //Exit beta search loop if truncated search is already a success. 
                    if(intSearchType==1) {
						//Re-screen failed, so check to determine whether all of the alpha graph nodes have been used
                        if(intUsedAlphaAtoms == ptrAlphaIDNode->intNumberAtoms) {
                            //Return complete successful search.
                            intSearchResult=1;
							break;
                        }
					}
               }//Closure of QK_Screen "if block"

          }//End of occupancy screen "if block"

          //Flag as new subfragment
          intNewSubFrag=1;
          
          //Block to determine how to proceed with the beta subfagment file.
          if(intQueryType) {
               //Continue screen/search in current bin (pocket)
               if(ptrBetaIDNode->ptrNextSubFrag != ptrBetaIDList->ptrBinSortLoc[intAlphaBinLoc+1]) {
                    ptrBetaIDNode=ptrBetaIDNode->ptrNextSubFrag;
               }
               //Determine next appropriate bin (pocket)
               else {
                    intAlphaBinLoc++;
                    //Increment through bin array until a location is found in the alpha occupation key, while the counter is less than the number of pockets, and the bin pointer is not null
                    while((!(ptrAlphaIDNode->intOccupancyKey & (1<<intAlphaBinLoc)) && (intAlphaBinLoc < intNumScreenPockets)) || (ptrBetaIDList->ptrBinSortLoc[intAlphaBinLoc+1] == ptrBetaIDList->ptrBinSortLoc[intAlphaBinLoc])) {
                         intAlphaBinLoc++;
                    }
                    
                    //If an appropriate location was found.
                    if(intAlphaBinLoc<intNumScreenPockets) ptrBetaIDNode=ptrBetaIDList->ptrBinSortLoc[intAlphaBinLoc];
                    //If no appropriate bin for this alpha was found.
                    else ptrBetaIDNode=0;
               }
          }
          //Just use next sequential subfrag if not a "file" search
          else ptrBetaIDNode=ptrBetaIDNode->ptrNextSubFrag;

     }//Closure of Beta search loop for given alpha graph

     //If enumerating search option is selected and the intersection (bitwise "and") is 
     //equivalent to the complete alpha graph, then perform the enumerating, exact cover alogrithm (Kreher, 1997)
     if(intSearchType==2) {
          
          //Perform exact cover if EC_Check covers the entire query molecule.
		 if(ptrAlphaIDNode->Compare_EC_Check()) {
               
               //If at least one subfragment grouping was detected, then signal a success.
               if(KreherExactCover(ptrAlphaIDNode,intSF_ID,intSF_Quant)) intSuccessSearch=1;
               //If no subfragment grouping was found, signal a failure.
               else intSuccessSearch=0;
          }
     }
     
     //Check to determine whether to continue in the alpha loop or to exit loop
     if(intQueryType) {
          
		  //Call "results" determination function for proper return value in a single query search.
		  MO_Return_Values(intSearchType,intSuccessSearch,intSearchResult);

		  //Output the results for the current query in the query file if it was a successful search.
		  MO_File_Output(ptrAlphaIDNode,intSearchResult,OutputAlphaFile,intSF_ID,intSF_Quant,intMF_ID,intMF_Quant);

		  //Check for end of file. If not, then read in next fragment.          
          if(InputAlphaFile.eof()) {
               InputAlphaFile.close();
			   OutputAlphaFile.close();
               break;
          }

          delete ptrAlphaIDNode;        
          ptrAlphaIDNode=0;
          
          //Create new alpha identity node.
          ptrAlphaIDNode=new ChemComboID;
          
          //Input the new query string.
		  InputAlphaFile>>ptrTmpAlphaString;
		  
		  //Place recently read subfragment string into new alpha identity node.
          ptrAlphaIDNode->chrChemSyntaxString=new char[strlen(ptrTmpAlphaString)+1];
		  strcpy(ptrAlphaIDNode->chrChemSyntaxString,ptrTmpAlphaString);
		  
          //Clear fragment detection arrays for next alpha structure
          for(j=0;j<100;j++) {
               intSF_ID[j]=0;
			   intSF_Quant[j]=0;
          }

		  //Clear molecular feature detection arrays for next alpha structure
          for(j=0;j<intMF_Detect;j++) {
               intMF_ID[j]=0;
			   intMF_Quant[j]=0;
          }

          //Reset the temporary quantity key
		  ptrTemp_Alpha_QK.Reset_QK();
		  
		  //Reset hash table for next alpha structure.
          if(ptrHash_Table) ptrHash_Table->Reset_HT();
     }
     //Exit Alpha loop
	 else break;
     
//Closure of Alpha while loop
}

//Cleanup dynamic structures prior to exiting
if(ptrAlphaIDNode) delete ptrAlphaIDNode;
ptrAlphaIDNode=0;
if(ptrMS) delete ptrMS;
ptrMS=0;
if(ptrBetaIDList) delete ptrBetaIDList;
ptrBetaIDList=0;
if(ptrHash_Table) delete ptrHash_Table;
ptrHash_Table=0;
if(ptrTmpAlphaString) delete ptrTmpAlphaString;
ptrTmpAlphaString=0;

//Call "results" determination function for proper return value in a single query search.
if(intQueryType==0) MO_Return_Values(intSearchType,intSuccessSearch,intSearchResult);
//Automatically return 0 when a file query is performed.
else intSearchResult=0;

return;

}




