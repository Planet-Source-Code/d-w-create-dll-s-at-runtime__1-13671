Attribute VB_Name = "CPU_Data"
Option Explicit

Private Declare Function GetProcessor Lib "getcpu.dll" ( _
  ByVal strCpu As String, ByVal strVendor As String, ByVal strVendorString As String, _
  StructuraCacheL2 As STRUCT_L1CACHE, StructuraCacheL1I As STRUCT_L1CACHE, _
  StructuraCacheL1D As STRUCT_L1CACHE, StructuraTLB As STRUCT_L1CACHETLB, _
  StructuraCacheL3 As STRUCT_L1CACHE, StructuraFMS As STRUCT_FMS, ByVal sSpeed As String) As Long
Private Declare Function GetProcessorFeatures Lib "getcpu.dll" ( _
lStructCpuFeat As STRUCT_CPU_FEAT) As Long

Private Type STRUCT_CPU_FEAT
  lCPUID As Long      'CPU identification routine support
  lCPUID_STD As Long  'CPUID standard identification routine support
  lCPUID_EXT As Long  'CPUID extended support
  lTSC As Long        'Time Stamp Counter support
  lMMX As Long        'MMX support - I guess everybody knows what's this
  lCMOV As Long       'Conditional Move Instruction support
  l3DNOW As Long      '3DNOW support - AMD
  l3DNOW_EXT As Long  '3DNOW extended support - for Sharptooth, for instance
  lSSE_MMX As Long    'Streaming SIMD Extensions 1 - SSE1 (for P3 mainly)
  lSSE_FPU As Long    'SSE 2
  lK6_MTRR As Long    'K6 Memory Type Registers
  lP6_MTRR As Long    'P6 Memory Type Registers
End Type

Private Type STRUCT_L1CACHE
  dwSize As Long        'size, in kb
  dwAssoc As Long       'ways associative
  dwLineSize As Long    'line size
  dwLinesPerTag As Long 'lines per tag
  dwPages As Long
End Type
Private Type STRUCT_L1CACHETLB
  dwAssocI As Long    '-way associative for level 1 cache instructions
  dwEntriesI As Long  'number of entries
  dwAssocD As Long    'identically, for data
  dwEntriesD As Long  'the same
  dwPagesI As Long
  dwPagesD As Long
End Type

Private Type STRUCT_FMS
  dwFamily As Long
  dwModel As Long
  dwStepping As Long
End Type

Private structL2Cache As STRUCT_L1CACHE
Private structL1CacheInstructions As STRUCT_L1CACHE
Private structL1CacheData As STRUCT_L1CACHE
Private structL1CacheTLB As STRUCT_L1CACHETLB
Private structL3Cache As STRUCT_L1CACHE
Private lStruct As STRUCT_CPU_FEAT


Public Function CPUData() As String

Dim structFMS As STRUCT_FMS
Dim sCpu As String * 255
Dim sVendor As String * 255
Dim sVendorString As String * 255
Dim sL2Cache As String * 255
Dim strSpeed As String * 255

Dim CPUType As String
Dim CPUFamily As String
Dim CPUModel As String
Dim CPUStepping As String
Dim Vendor As String
Dim VendorString As String
Dim CPUSpeed As String


GetProcessor sCpu, sVendor, sVendorString, structL2Cache, _
structL1CacheInstructions, structL1CacheData, structL1CacheTLB, _
structL3Cache, structFMS, strSpeed

GetProcessorFeatures lStruct

CPUType = StripNull(sCpu)
CPUFamily = structFMS.dwFamily
CPUModel = structFMS.dwModel
CPUStepping = structFMS.dwStepping
Vendor = StripNull(sVendor)
VendorString = StripNull(sVendorString)
CPUSpeed = StripNull(strSpeed)

CPUData = Vendor & ": " & VendorString & ", type " & CPUType & _
 ", family " & CPUFamily & ", model " & CPUModel & ", stepping " & _
 CPUStepping & ", speed " & CPUSpeed & " Mhz. " & CPUFeatures
End Function


Private Function StripNull(sInput As String) As String
Dim nPos As Integer
nPos = InStr(sInput, Chr(0))
If nPos <> 0 Then
  StripNull = Left(sInput, nPos - 1)
Else
  StripNull = sInput
End If
End Function




Private Function CPUFeatures() As String
Dim CPUID As Boolean
Dim CPUID_STD As Boolean
Dim CPUID_EXT As Boolean
Dim TSC As Boolean
Dim MMX As Boolean
Dim CMOV As Boolean
Dim TriDNow As Boolean
Dim TriDNow_EXT As Boolean
Dim SSEMMX As Boolean
Dim SSEFPU As Boolean
Dim KSixMTRR As Boolean
Dim PSixMTRR As Boolean
Dim sTemp As String
CPUID = CBool(lStruct.lCPUID)
CPUID_STD = CBool(lStruct.lCPUID_STD)
CPUID_EXT = CBool(lStruct.l3DNOW_EXT)
TSC = CBool(lStruct.lTSC)
MMX = CBool(lStruct.lMMX)
CMOV = CBool(lStruct.lCMOV)
TriDNow = CBool(lStruct.l3DNOW)
TriDNow_EXT = CBool(lStruct.l3DNOW_EXT)
SSEMMX = CBool(lStruct.lSSE_MMX)
SSEFPU = CBool(lStruct.lSSE_FPU)
KSixMTRR = CBool(lStruct.lK6_MTRR)
PSixMTRR = CBool(lStruct.lP6_MTRR)
sTemp = " Features include: "
If CPUID Then
sTemp = sTemp & "CPU ID support"
End If
If CPUID_STD Then
sTemp = sTemp & ", " & "CPU  Standard ID support"
End If
If CPUID_EXT Then
sTemp = sTemp & ", " & "CPU Extended ID support"
End If
If TSC Then
sTemp = sTemp & ", " & "Time Stamp Coding support"
End If
If MMX Then
sTemp = sTemp & ", " & "MMX instruction support"
End If
If CMOV Then
sTemp = sTemp & ", " & "Conditional Move Instruction support"
End If
If TriDNow Then
sTemp = sTemp & ", " & "3D-NOW! support"
End If
If TriDNow_EXT Then
sTemp = sTemp & ", " & "Extended 3D-NOW! support"
End If
If SSEMMX Then
sTemp = sTemp & ", " & "Streaming SIMD Extensions 1"
End If
If SSEFPU Then
sTemp = sTemp & ", " & "Streaming SIMD Extensions 2"
End If
If KSixMTRR Then
sTemp = sTemp & ", " & "K6 Memory Type Registers"
End If
If PSixMTRR Then
sTemp = sTemp & ", " & "P6 Memory Type Registers"
End If
sTemp = sTemp & "."
If sTemp = " Features include: ." Then
sTemp = ""
End If
CPUFeatures = sTemp
End Function


Private Function LOneCacheData() As String
If structL1CacheData.dwSize <> 0 Then
LOneCacheData = "L1 Cache Data:" & vbTab & vbTab & vbTab & structL1CacheData.dwSize & " kb, " & _
          structL1CacheData.dwAssoc & "-way associative, " & structL1CacheData.dwLineSize & _
          " byte line size, " & structL1CacheData.dwLinesPerTag & " lines per tag"
Else
LOneCacheData = ""
End If
End Function


Private Function LOneCacheInstructions() As String
If structL1CacheInstructions.dwSize <> 0 Then
LOneCacheInstructions = "L1 Cache Instructions:" & vbTab & vbTab & structL1CacheInstructions.dwSize & " kb, " & _
    structL1CacheInstructions.dwAssoc & "-way associative, " & structL1CacheInstructions.dwLineSize & _
    " byte line size, " & structL1CacheInstructions.dwLinesPerTag & " lines per tag"
Else
LOneCacheInstructions = ""
End If
End Function


Private Function LOneDataTLB() As String
If structL1CacheTLB.dwAssocD > 0 Then
LOneDataTLB = "L1 Data TLB:" & vbTab & vbTab & vbTab & structL1CacheTLB.dwAssocD & "-way set associative, " & _
    structL1CacheTLB.dwEntriesD & " entries"
Else
LOneDataTLB = ""
End If
End Function

Private Function LOneInstTLB() As String
If structL1CacheTLB.dwAssocI <> 0 Then
LOneInstTLB = "L1 Instruction TLB:" & vbTab & vbTab & structL1CacheTLB.dwAssocI & "-way set associative, " & _
          structL1CacheTLB.dwEntriesI & " entries"
Else
LOneInstTLB = ""
End If
End Function


Private Function LThreeCacheSize() As String
If structL3Cache.dwSize > 0 Then
LThreeCacheSize = "L3 Cache Size" & vbTab & vbTab & vbTab & structL3Cache.dwSize & " kb, " & _
    structL3Cache.dwAssoc & "-way set associative, " & structL3Cache.dwLineSize & _
    " byte line size"
Else
LThreeCacheSize = ""
End If
End Function

Private Function LTwoCacheSize() As String
If structL2Cache.dwSize > 0 Then
LTwoCacheSize = "L2 Cache Size" & vbTab & vbTab & vbTab & structL2Cache.dwSize & " kb, " & _
          structL2Cache.dwAssoc & "-way set associative, " & structL2Cache.dwLineSize & _
          " byte line size"
Else
LTwoCacheSize = ""
End If
End Function


