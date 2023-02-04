Attribute VB_Name = "xlAppScript_setup"
'/\______________________________________________________________________________________________________________________
'//
'//     xlAppScript Setup
'/\_____________________________________________________________________________________________________________________________
'//
'//     License Information:
'//
'//     Copyright (C) 2022-present, Autokit Technology.
'//
'//     Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'//
'//     1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'//
'//     2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'//
'//     3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
'//
'//     THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
'//     THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'//     (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
'//     HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'//
'/\_____________________________________________________________________________________________________________________________
'//
'//     Latest Revision: 1/31/2023
'/\_____________________________________________________________________________________________________________________________
'//
'//     Developer(s): anz7re (André)
'//     Contact: support@xlappscript.org | support@autokit.tech | anz7re@autokit.tech
'//     Web: xlappscript.org | autokit.tech
'/\_____________________________________________________________________________________________________________________________

Public Function connectWb()

'/\__________________________________________________________________________________________
'//
'//     A function for setting up a runtime environment (workbook) to interact w/ xlAppScript
'/\__________________________________________________________________________________________

On Error GoTo ErrEnd

nl = vbNewLine

'//Label script memory addresses (Very important!)

'//Multi-memory block specific addresses
Range("MAA1").name = "xlasKinLabelMod"
Range("MAB1").name = "xlasKinValueMod"
Range("MAC1").name = "xlasKinLabel"
Range("MAD1").name = "xlasKinValue"
Range("MAE1").name = "xlasState"
Range("MAF1").name = "xlasArticle"
Range("MAG1").name = "xlasGroup"
Range("MAH1").name = "xlasList" '(3/15/2022)
Range("MAL1").name = "xlasLib"
'//Single-memory block specific addresses
Range("MAS1").name = "xlasAppLoad" 'Autokit applications(2/9/2022)
Range("MAS2").name = "xlasEnvironment"
Range("MAS3").name = "xlasBlock" '(2/28/2022)
Range("MAS4").name = "xlasGoto"
Range("MAS5").name = "xlasInputField" '8/1/2022
Range("MAS6").name = "xlasInvert" 'ctrl box+
Range("MAS7").name = "xlasKeyCtrl" 'Autokit applications
Range("MAS8").name = "xlasRemember" 'ctrl box+
Range("MAS9").name = "xlasConsoleType" 'ctrl box+
Range("MAS10").name = "xlasAMemory" 'ctrl box+
Range("MAS11").name = "xlasSaveFile" 'ctrl box+
Range("MAS12").name = "xlasSilent" 'Autokit applications
Range("MAS13").name = "xlasCtrlBoxFColor" 'ctrl box+
Range("MAS14").name = "xlasCtrlBoxBColor" 'ctrl box+
Range("MAS15").name = "xlasGlobalControl" '4/18/2022
Range("MAS16").name = "xlasLocalContain" '(3/17/2022)
Range("MAS17").name = "xlasLocalStatic" '(3/13/2022)
Range("MAS18").name = "xlasUpdateEnable" '(2/24/2022)
Range("MAS19").name = "xlasWinForm" 'Autokit applications
Range("MAS20").name = "xlasWinFormLast" 'Autokit applications
Range("MAS21").name = "xlasWinFormX" '5/12/2022
Range("MAS22").name = "xlasWinFormY" '5/12/2022
Range("MAS23").name = "xlasLibCount" '1/30/2023
Range("MAS24").name = "xlasLibErrLvl" '4/22/2022
Range("MAS25").name = "xlasErrRef"
Range("MAS26").name = "xlasEnd"
Range("MAS27").name = "xlasLink": Range("xlasLink").Value = 1
Range("MAS79").name = "xlasBlkAddr79" '3/17/2022
Range("MAS80").name = "xlasBlkAddr80" '3/17/2022
Range("MAS81").name = "xlasBlkAddr81" '3/17/2022
Range("MAS82").name = "xlasBlkAddr82" '3/17/2022
Range("MAS83").name = "xlasBlkAddr83" '3/17/2022
Range("MAS84").name = "xlasBlkAddr84" '3/17/2022
Range("MAS85").name = "xlasBlkAddr85" '3/17/2022
Range("MAS86").name = "xlasBlkAddr86" '3/17/2022
Range("MAS87").name = "xlasBlkAddr87" '3/17/2022
Range("MAS88").name = "xlasBlkAddr88" '3/17/2022
Range("MAS89").name = "xlasBlkAddr89" '3/17/2022
Range("MAS90").name = "xlasBlkAddr90" '3/17/2022
Range("MAS91").name = "xlasBlkAddr91" '3/17/2022
Range("MAS92").name = "xlasBlkAddr92" '3/17/2022
Range("MAS93").name = "xlasBlkAddr93" '3/17/2022
Range("MAS94").name = "xlasBlkAddr94" '3/17/2022
Range("MAS95").name = "xlasBlkAddr95" '3/17/2022
Range("MAS96").name = "xlasBlkAddr96" '3/17/2022
Range("MAS97").name = "xlasBlkAddr97" '3/17/2022
Range("MAS98").name = "xlasBlkAddr98" '3/17/2022
Range("MAS99").name = "xlasBlkAddr99" '3/17/2022
Range("MAS100").name = "xlasBlkAddr100" '4/21/2022
Range("MAS101").name = "xlasBlkAddr101" '4/21/2022
Range("MAS102").name = "xlasBlkAddr102" '4/21/2022
Range("MAS103").name = "xlasBlkAddr103" '4/21/2022
Range("MAS104").name = "xlasBlkAddr104" '4/21/2022
Range("MAS105").name = "xlasBlkAddr105" '4/21/2022
Range("MAS106").name = "xlasBlkAddr106" '4/21/2022
Range("MAS107").name = "xlasBlkAddr107" '4/21/2022
Range("MAS108").name = "xlasBlkAddr108" '4/21/2022
Range("MAS109").name = "xlasBlkAddr109" '4/21/2022
Range("MAS110").name = "xlasBlkAddr110" '4/21/2022
Range("MAS111").name = "xlasBlkAddr111" '4/21/2022
Range("MAS112").name = "xlasBlkAddr112" '4/21/2022
Range("MAS113").name = "xlasBlkAddr113" '4/21/2022
Range("MAS114").name = "xlasBlkAddr114" '4/21/2022
Range("MAS115").name = "xlasBlkAddr115" '4/21/2022
Range("MAS116").name = "xlasBlkAddr116" '4/21/2022
Range("MAS117").name = "xlasBlkAddr117" '4/21/2022
Range("MAS118").name = "xlasBlkAddr118" '4/21/2022
Range("MAS119").name = "xlasBlkAddr119" '4/21/2022
Range("MAS120").name = "xlasBlkAddr120" '4/21/2022
Range("MAS121").name = "xlasBlkAddr121" '4/21/2022
Range("MAS122").name = "xlasBlkAddr122" '4/21/2022
Range("MAS123").name = "xlasBlkAddr123" '4/21/2022
Range("MAS124").name = "xlasBlkAddr124" '4/21/2022
Range("MAS125").name = "xlasBlkAddr125" '4/21/2022
Range("MAS126").name = "xlasBlkAddr126" '4/21/2022
Range("MAS127").name = "xlasBlkAddr127" '4/21/2022
Range("MAS128").name = "xlasBlkAddr128" '4/21/2022
Range("MAS129").name = "xlasBlkAddr129" '4/21/2022
Range("MAS130").name = "xlasBlkAddr130" '4/21/2022
Range("MAS131").name = "xlasBlkAddr131" '4/21/2022
Range("MAS132").name = "xlasBlkAddr132" '4/21/2022
Range("MAS133").name = "xlasBlkAddr133" '4/21/2022
Range("MAS134").name = "xlasBlkAddr134" '4/21/2022
Range("MAS135").name = "xlasBlkAddr135" '4/21/2022
Range("MAS136").name = "xlasBlkAddr136" '4/21/2022
Range("MAS137").name = "xlasBlkAddr137" '4/21/2022
Range("MAS138").name = "xlasBlkAddr138" '4/21/2022
Range("MAS139").name = "xlasBlkAddr139" '4/21/2022
Range("MAS140").name = "xlasBlkAddr140" '4/21/2022
Range("MAS141").name = "xlasBlkAddr141" '4/21/2022
Range("MAS142").name = "xlasBlkAddr142" '4/21/2022
Range("MAS143").name = "xlasBlkAddr143" '4/21/2022
Range("MAS144").name = "xlasBlkAddr144" '4/21/2022
Range("MAS145").name = "xlasBlkAddr145" '4/21/2022
Range("MAS146").name = "xlasBlkAddr146" '4/21/2022
Range("MAS147").name = "xlasBlkAddr147" '4/21/2022
Range("MAS148").name = "xlasBlkAddr148" '4/21/2022
Range("MAS149").name = "xlasBlkAddr149" '4/21/2022
Range("MAS150").name = "xlasBlkAddr150" '4/21/2022
Range("MAS151").name = "xlasBlkAddr151" '4/21/2022
Range("MAS152").name = "xlasBlkAddr152" '4/21/2022
Range("MAS153").name = "xlasBlkAddr153" '4/21/2022
Range("MAS154").name = "xlasBlkAddr154" '4/21/2022
Range("MAS155").name = "xlasBlkAddr155" '4/21/2022
Range("MAS156").name = "xlasBlkAddr156" '4/21/2022
Range("MAS157").name = "xlasBlkAddr157" '4/21/2022
Range("MAS158").name = "xlasBlkAddr158" '4/21/2022
Range("MAS159").name = "xlasBlkAddr159" '4/21/2022
Range("MAS160").name = "xlasBlkAddr160" '4/21/2022
Range("MAS161").name = "xlasBlkAddr161" '4/21/2022
Range("MAS162").name = "xlasBlkAddr162" '4/21/2022
Range("MAS163").name = "xlasBlkAddr163" '4/21/2022
Range("MAS164").name = "xlasBlkAddr164" '4/21/2022
Range("MAS165").name = "xlasBlkAddr165" '4/21/2022
Range("MAS166").name = "xlasBlkAddr166" '4/21/2022
Range("MAS167").name = "xlasBlkAddr167" '4/21/2022
Range("MAS168").name = "xlasBlkAddr168" '4/21/2022
Range("MAS169").name = "xlasBlkAddr169" '4/21/2022
Range("MAS170").name = "xlasBlkAddr170" '4/21/2022
Range("MAS171").name = "xlasBlkAddr171" '4/21/2022
Range("MAS172").name = "xlasBlkAddr172" '4/21/2022
Range("MAS173").name = "xlasBlkAddr173" '4/21/2022
Range("MAS174").name = "xlasBlkAddr174" '4/21/2022
Range("MAS175").name = "xlasBlkAddr175" '4/21/2022
Range("MAS176").name = "xlasBlkAddr176" '4/21/2022
Range("MAS177").name = "xlasBlkAddr177" '4/21/2022
Range("MAS178").name = "xlasBlkAddr178" '4/21/2022
Range("MAS179").name = "xlasBlkAddr179" '4/21/2022
Range("MAS180").name = "xlasBlkAddr180" '4/21/2022
Range("MAS181").name = "xlasBlkAddr181" '4/21/2022
Range("MAS182").name = "xlasBlkAddr182" '4/21/2022
Range("MAS183").name = "xlasBlkAddr183" '4/21/2022
Range("MAS184").name = "xlasBlkAddr184" '4/21/2022
Range("MAS185").name = "xlasBlkAddr185" '4/21/2022
Range("MAS186").name = "xlasBlkAddr186" '4/21/2022
Range("MAS187").name = "xlasBlkAddr187" '4/21/2022
Range("MAS188").name = "xlasBlkAddr188" '4/21/2022
Range("MAS189").name = "xlasBlkAddr189" '4/21/2022
Range("MAS190").name = "xlasBlkAddr190" '4/21/2022
Range("MAS191").name = "xlasBlkAddr191" '4/21/2022
Range("MAS192").name = "xlasBlkAddr192" '4/21/2022
Range("MAS193").name = "xlasBlkAddr193" '4/21/2022
Range("MAS194").name = "xlasBlkAddr194" '4/21/2022
Range("MAS195").name = "xlasBlkAddr195" '4/21/2022
Range("MAS196").name = "xlasBlkAddr196" '4/21/2022
Range("MAS197").name = "xlasBlkAddr197" '4/21/2022
Range("MAS198").name = "xlasBlkAddr198" '4/21/2022
Range("MAS199").name = "xlasBlkAddr199" '4/21/2022
Range("MAS200").name = "xlasBlkAddr200" '4/21/2022
Range("MAS201").name = "xlasBlkAddr201" '4/21/2022
Range("MAS202").name = "xlasBlkAddr202" '4/21/2022
Range("MAS203").name = "xlasBlkAddr203" '4/21/2022
Range("MAS204").name = "xlasBlkAddr204" '4/21/2022
Range("MAS205").name = "xlasBlkAddr205" '4/21/2022
Range("MAS206").name = "xlasBlkAddr206" '4/21/2022
Range("MAS207").name = "xlasBlkAddr207" '4/21/2022
Range("MAS208").name = "xlasBlkAddr208" '4/21/2022
Range("MAS209").name = "xlasBlkAddr209" '4/21/2022
Range("MAS210").name = "xlasBlkAddr210" '4/21/2022
Range("MAS211").name = "xlasBlkAddr211" '4/21/2022
Range("MAS212").name = "xlasBlkAddr212" '4/21/2022
Range("MAS213").name = "xlasBlkAddr213" '4/21/2022
Range("MAS214").name = "xlasBlkAddr214" '4/21/2022
Range("MAS215").name = "xlasBlkAddr215" '4/21/2022
Range("MAS216").name = "xlasBlkAddr216" '4/21/2022
Range("MAS217").name = "xlasBlkAddr217" '4/21/2022
Range("MAS218").name = "xlasBlkAddr218" '4/21/2022
Range("MAS219").name = "xlasBlkAddr219" '4/21/2022
Range("MAS220").name = "xlasBlkAddr220" '4/21/2022
Range("MAS221").name = "xlasBlkAddr221" '4/21/2022
Range("MAS222").name = "xlasBlkAddr222" '4/21/2022
Range("MAS223").name = "xlasBlkAddr223" '4/21/2022
Range("MAS224").name = "xlasBlkAddr224" '4/21/2022
Range("MAS225").name = "xlasBlkAddr225" '4/21/2022
Range("MAS226").name = "xlasBlkAddr226" '4/21/2022
Range("MAS227").name = "xlasBlkAddr227" '4/21/2022
Range("MAS228").name = "xlasBlkAddr228" '4/21/2022
Range("MAS229").name = "xlasBlkAddr229" '4/21/2022
Range("MAS230").name = "xlasBlkAddr230" '4/21/2022
Range("MAS231").name = "xlasBlkAddr231" '4/21/2022
Range("MAS232").name = "xlasBlkAddr232" '4/21/2022
Range("MAS233").name = "xlasBlkAddr233" '4/21/2022
Range("MAS234").name = "xlasBlkAddr234" '4/21/2022
Range("MAS235").name = "xlasBlkAddr235" '4/21/2022
Range("MAS236").name = "xlasBlkAddr236" '4/21/2022
Range("MAS237").name = "xlasBlkAddr237" '4/21/2022
Range("MAS238").name = "xlasBlkAddr238" '4/21/2022
Range("MAS239").name = "xlasBlkAddr239" '4/21/2022
Range("MAS240").name = "xlasBlkAddr240" '4/21/2022
Range("MAS241").name = "xlasBlkAddr241" '4/21/2022
Range("MAS242").name = "xlasBlkAddr242" '4/21/2022
Range("MAS243").name = "xlasBlkAddr243" '4/21/2022
Range("MAS244").name = "xlasBlkAddr244" '4/21/2022
Range("MAS245").name = "xlasBlkAddr245" '4/21/2022
Range("MAS246").name = "xlasBlkAddr246" '4/21/2022
Range("MAS247").name = "xlasBlkAddr247" '4/21/2022
Range("MAS248").name = "xlasBlkAddr248" '4/21/2022
Range("MAS249").name = "xlasBlkAddr249" '4/21/2022
Range("MAS250").name = "xlasBlkAddr250" '4/21/2022
Range("MAS251").name = "xlasBlkAddr251" '4/21/2022
Range("MAS252").name = "xlasBlkAddr252" '4/21/2022
Range("MAS253").name = "xlasBlkAddr253" '4/21/2022
Range("MAS254").name = "xlasBlkAddr254" '4/21/2022
Range("MAS255").name = "xlasBlkAddr255" '4/21/2022
Range("MAS256").name = "xlasBlkAddr256" '4/21/2022
Range("MAS257").name = "xlasBlkAddr257" '4/21/2022
Range("MAS258").name = "xlasBlkAddr258" '4/21/2022
Range("MAS259").name = "xlasBlkAddr259" '4/21/2022
Range("MAS260").name = "xlasBlkAddr260" '4/21/2022
Range("MAS261").name = "xlasBlkAddr261" '4/21/2022
Range("MAS262").name = "xlasBlkAddr262" '4/21/2022
Range("MAS263").name = "xlasBlkAddr263" '4/21/2022
Range("MAS264").name = "xlasBlkAddr264" '4/21/2022
Range("MAS265").name = "xlasBlkAddr265" '4/21/2022
Range("MAS266").name = "xlasBlkAddr266" '4/21/2022
Range("MAS267").name = "xlasBlkAddr267" '4/21/2022
Range("MAS268").name = "xlasBlkAddr268" '4/21/2022
Range("MAS269").name = "xlasBlkAddr269" '4/21/2022
Range("MAS270").name = "xlasBlkAddr270" '4/21/2022
Range("MAS271").name = "xlasBlkAddr271" '4/21/2022
Range("MAS272").name = "xlasBlkAddr272" '4/21/2022
Range("MAS273").name = "xlasBlkAddr273" '4/21/2022
Range("MAS274").name = "xlasBlkAddr274" '4/21/2022
Range("MAS275").name = "xlasBlkAddr275" '4/21/2022
Range("MAS276").name = "xlasBlkAddr276" '4/21/2022
Range("MAS277").name = "xlasBlkAddr277" '4/21/2022
Range("MAS278").name = "xlasBlkAddr278" '4/21/2022
Range("MAS279").name = "xlasBlkAddr279" '4/21/2022

'//Create target script locations
If Dir(drv & envHome & "\.z7", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7")
If Dir(drv & envHome & "\.z7\utility", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility")
If Dir(drv & envHome & "\.z7\utility\debug", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\debug")
If Dir(drv & envHome & "\.z7\utility\temp", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\temp")
If Dir(drv & envHome & "\.z7\utility\miss", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\miss")
If Dir(drv & envHome & "\.z7\utility\miss\colors", vbDirectory) = "" Then MkDir (drv & envHome & "\.z7\utility\miss\colors")

MsgBox ("xlAppScript runtime environment connection is complete." & nl & nl & _
"Current environment: " & ThisWorkbook.name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbInformation

Exit Function

ErrEnd:
MsgBox ("There was an issue trying to connect this runtime environment." & nl & nl & _
"Current environment: " & ThisWorkbook.name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbCritical

End Function
Public Function disconnectWb()

'/\_____________________________________________________________________________________
'//
'//     A function for removing an environment from interacting w/ xlAppScript
'/\_____________________________________________________________________________________


On Error GoTo ErrEnd

nl = vbNewLine

'//Remove script addresses

'//Multi-memory block specific addresses
ActiveWorkbook.Names("xlasKinLabelMod").Delete
ActiveWorkbook.Names("xlasKinValueMod").Delete
ActiveWorkbook.Names("xlasKinLabel").Delete
ActiveWorkbook.Names("xlasKinValue").Delete
ActiveWorkbook.Names("xlasState").Delete
ActiveWorkbook.Names("xlasArticle").Delete
ActiveWorkbook.Names("xlasGroup").Delete
ActiveWorkbook.Names("xlasList").Delete '(3/15/2022)
ActiveWorkbook.Names("xlasLib").Delete
'//Single-memory block specific addresses
ActiveWorkbook.Names("xlasAppLoad").Delete 'Autokit applications(2/9/2022)
ActiveWorkbook.Names("xlasEnvironment").Delete
ActiveWorkbook.Names("xlasBlock").Delete '(2/28/2022)
ActiveWorkbook.Names("xlasGoto").Delete
ActiveWorkbook.Names("xlasInputField").Delete '8/1/2022
ActiveWorkbook.Names("xlasInvert").Delete 'ctrl box+
ActiveWorkbook.Names("xlasKeyCtrl").Delete 'Autokit applications
ActiveWorkbook.Names("xlasSilent").Delete 'Autokit applications
ActiveWorkbook.Names("xlasRemember").Delete 'ctrl box+
ActiveWorkbook.Names("xlasConsoleType").Delete 'ctrl box+
ActiveWorkbook.Names("xlasAMemory").Delete 'ctrl box+
ActiveWorkbook.Names("xlasSaveFile").Delete 'ctrl box+
ActiveWorkbook.Names("xlasCtrlBoxBColor").Delete 'ctrl box+
ActiveWorkbook.Names("xlasCtrlBoxFColor").Delete 'ctrl box+
ActiveWorkbook.Names("xlasGlobalControl").Delete '4/18/2022
ActiveWorkbook.Names("xlasLocalContain").Delete '(3/17/2022)
ActiveWorkbook.Names("xlasLocalStatic").Delete '(3/13/2022)
ActiveWorkbook.Names("xlasUpdateEnable").Delete '(2/24/2022)
ActiveWorkbook.Names("xlasWinForm").Delete 'Autokit applications
ActiveWorkbook.Names("xlasWinFormLast").Delete 'Autokit applications
ActiveWorkbook.Names("xlasWinFormX").Delete '5/12/2022
ActiveWorkbook.Names("xlasWinFormY").Delete '5/12/2022
ActiveWorkbook.Names("xlasLibCount").Delete '1/30/2023
ActiveWorkbook.Names("xlasLibErrLvl").Delete '4/22/2022
ActiveWorkbook.Names("xlasErrRef").Delete
ActiveWorkbook.Names("xlasEnd").Delete
Range("xlasLink").Value = vbNullString: ActiveWorkbook.Names("xlasLink").Delete
ActiveWorkbook.Names("xlasBlkAddr79").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr80").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr81").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr82").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr83").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr84").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr85").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr86").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr87").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr88").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr89").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr90").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr91").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr92").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr93").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr94").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr95").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr96").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr97").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr98").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr99").Delete '3/17/2022
ActiveWorkbook.Names("xlasBlkAddr100").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr101").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr102").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr103").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr104").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr105").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr106").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr107").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr108").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr109").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr110").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr111").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr112").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr113").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr114").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr115").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr116").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr117").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr118").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr119").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr120").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr121").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr122").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr123").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr124").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr125").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr126").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr127").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr128").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr129").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr130").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr131").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr132").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr133").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr134").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr135").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr136").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr137").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr138").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr139").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr140").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr141").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr142").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr143").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr144").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr145").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr146").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr147").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr148").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr149").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr150").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr151").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr152").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr153").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr154").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr155").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr156").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr157").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr158").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr159").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr160").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr161").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr162").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr163").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr164").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr165").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr166").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr167").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr168").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr169").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr170").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr171").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr172").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr173").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr174").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr175").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr176").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr177").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr178").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr179").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr180").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr181").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr182").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr183").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr184").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr185").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr186").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr187").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr188").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr189").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr190").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr191").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr192").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr193").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr194").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr195").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr196").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr197").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr198").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr199").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr200").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr201").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr202").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr203").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr204").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr205").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr206").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr207").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr208").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr209").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr210").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr211").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr212").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr213").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr214").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr215").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr216").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr217").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr218").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr219").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr220").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr221").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr222").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr223").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr224").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr225").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr226").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr227").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr228").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr229").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr230").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr231").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr232").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr233").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr234").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr235").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr236").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr237").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr238").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr239").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr240").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr241").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr242").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr243").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr244").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr245").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr246").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr247").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr248").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr249").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr250").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr251").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr252").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr253").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr254").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr255").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr256").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr257").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr258").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr259").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr260").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr261").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr262").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr263").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr264").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr265").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr266").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr267").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr268").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr269").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr270").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr271").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr272").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr273").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr274").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr275").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr276").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr277").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr278").Delete '4/21/2022
ActiveWorkbook.Names("xlasBlkAddr279").Delete '4/21/2022

MsgBox ("xlAppScript runtime environment disconnection is complete." & nl & nl & _
"Current environment: " & ThisWorkbook.name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbInformation

Exit Function

ErrEnd:
MsgBox ("There was an issue trying to disconnect this runtime environment." & nl & nl & _
"Current environment: " & ThisWorkbook.name & nl & nl & _
"Current environment path: " & ThisWorkbook.FullName), vbCritical

End Function

