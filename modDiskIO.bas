Attribute VB_Name = "modDiskIO"

Option Explicit

Private Const MAX_PATH = 260
Private Const MAX_DEVICES = 64
Private Const MAX_PARTITIONS = 32
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OBJ_CASE_INSENSITIVE = &H40
Private Const OBJ_KERNEL_HANDLE = &H200
Private Const SYNCHRONIZE = &H100000
Private Const FILE_READ_ATTRIBUTES = &H80
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_OPEN = &H1
Private Const FILE_SYNCHRONOUS_IO_NONALERT = &H20

Private Const FILE_DEVICE_DISK = &H7
Private Const IOCTL_DISK_BASE = FILE_DEVICE_DISK
Private Const METHOD_BUFFERED = 0
Private Const FILE_ANY_ACCESS = 0
Private Const FILE_READ_ACCESS = &H1

Private Const FILE_DEVICE_MASS_STORAGE = &H2D
Private Const IOCTL_STORAGE_BASE = FILE_DEVICE_MASS_STORAGE

Public Const MAX_DISK_INFO_PARAMS = 18
Public Const NO_DOS_LETTERS = "no DOS-like letters"

Public Enum DISK_INFO
    DriveIndex = 0 'sDiskInfo(0, UBound(sDiskInfo, 2)) = dev_num
    DrivePath = 1 'sDiskInfo(1, UBound(sDiskInfo, 2)) = sDev
    PartitionPath = 2 'sDiskInfo(2, UBound(sDiskInfo, 2)) = part_info
    PartitionLetter = 3 'sDiskInfo(3, UBound(sDiskInfo, 2)) = sTmp
    TotalMegaBytes = 4 'sDiskInfo(4, UBound(sDiskInfo, 2)) = lTotalMB
    UsedMegaBytes = 5 'sDiskInfo(5, UBound(sDiskInfo, 2)) = lUsedMB
    FreeMegaBytes = 6 'sDiskInfo(6, UBound(sDiskInfo, 2)) = lFreeMB
    AvailableMegaBytes = 7 'sDiskInfo(7, UBound(sDiskInfo, 2)) = lAvailMB
    VolumeName = 8 'sDiskInfo(8, UBound(sDiskInfo, 2)) = VolumeName
    VolumeFileSystem = 9 'sDiskInfo(9, UBound(sDiskInfo, 2)) = VolumeFS
    VolumeSerial = 10 'sDiskInfo(10, UBound(sDiskInfo, 2)) = VolSerial
    MatchedVolume = 11
    VolumeLettersAndFolders = 12
    PartitionNumber = 13
    PartitionStyle = 14 ' MBR, GPT, RAW
    PartitionGPT_Name = 15
    PartitionGPT_GUID = 16
    DriveBusType = 17
    DriveName = 18
End Enum

Private Enum PARTITION_STYLE
    PARTITION_STYLE_MBR = 0
    PARTITION_STYLE_GPT = 1
    PARTITION_STYLE_RAW = 2
End Enum

Private Enum STORAGE_BUS_TYPE
    BusTypeUnknown = &H0
    BusTypeScsi = &H1
    BusTypeAtapi = &H2
    BusTypeAta = &H3
    BusType1394 = &H4
    BusTypeSsa = &H5
    BusTypeFibre = &H6
    BusTypeUsb = &H7
    BusTypeRAID = &H8
    BusTypeiSCSI = &H9
    BusTypeSas = &HA
    BusTypeSata = &HB
    BusTypeMaxReserved = &H7F
    BusTypeAlignToLong = &HFFFFFFFF
End Enum

Private Enum STORAGE_PROPERTY_ID
  StorageDeviceProperty = &H0
  StorageAdapterProperty = &H1
  StorageDeviceIdProperty = &H2
  StorageDeviceUniqueIdProperty = &H3
  StorageDeviceWriteCacheProperty = &H4
  StorageMiniportProperty = &H5
  StorageAccessAlignmentProperty = &H6
  StorageDeviceSeekPenaltyProperty = &H7
  StorageDeviceTrimProperty = &H8
  StorageDeviceWriteAggregationProperty = &H9
  StorageDeviceDeviceTelemetryProperty = &HA
  StorageDeviceLBProvisioningProperty = &HB
  StorageDevicePowerProperty = &HC
  StorageDeviceCopyOffloadProperty = &HD
  StorageDeviceResiliencyProperty = &HE
  StorageDeviceAlignToLong = &HFFFFFFFF
End Enum

Private Enum STORAGE_QUERY_TYPE
  PropertyStandardQuery = &H0
  PropertyExistsQuery = &H1
  PropertyMaskQuery = &H2
  PropertyQueryMaxDefined = &H3
  PropertyQueryAlignToLong = &HFFFFFFFF
End Enum

Private Type STORAGE_PROPERTY_QUERY
    PropertyId As STORAGE_PROPERTY_ID
    QueryType As STORAGE_QUERY_TYPE
    AdditionalParameters(0 To 0) As Byte
End Type

Private Type STORAGE_DEVICE_DESCRIPTOR
    Version As Long
    Size As Long
    DeviceType As Byte
    DeviceTypeModifier As Byte
    RemovableMedia As Byte
    CommandQueueing As Byte
    VendorIdOffset As Long
    ProductIdOffset As Long
    ProductRevisionOffset As Long
    SerialNumberOffset As Long
    BusType As Long ' STORAGE_BUS_TYPE
    RawPropertiesLength As Long
    RawDeviceProperties(0 To 1000) As Byte
End Type

' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' NO SENCE IS THERE TO USING   LARGE_INTEGER   -  it is 8 bytes same as CURRENCY. And work with Currency is more slightly
' But it is needed for some DLL calls so let it be :)
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type UNICODE_STRING
    Length As Integer
    MaximumLength As Integer
    Buffer As Long
End Type

Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Private Type IO_STATUS_BLOCK
    Status As Long
    Information As Long
End Type

Private Type PARTITION_INFORMATION_MBR
    PartitionType As Byte
    BootIndicator As Byte
    RecognizedPartition As Byte
    UnusedByteForAlignTo4 As Byte
    HiddenSectors As Long
End Type

'typedef struct _PARTITION_INFORMATION_GPT {
'  GUID    PartitionType;
'  GUID    PartitionId;
'  DWORD64 Attributes;
'  WCHAR   Name[36];
'} PARTITION_INFORMATION_GPT, *PPARTITION_INFORMATION_GPT;

Private Type PARTITION_INFORMATION_GPT
    PartitionType As GUID
    PartitionId As GUID
    Attributes As Currency
    NameGPT(0 To 71) As Byte ' WideChar (doubled unicode chars) string with GPT partition name
End Type

'typedef struct {
'  PARTITION_STYLE PartitionStyle;
'  LARGE_INTEGER   StartingOffset;
'  LARGE_INTEGER   PartitionLength;
'  DWORD           PartitionNumber;
'  BOOLEAN         RewritePartition;
'  union {
'    PARTITION_INFORMATION_MBR Mbr;
'    PARTITION_INFORMATION_GPT Gpt;
'  };
'} PARTITION_INFORMATION_EX;

Private Type PARTITION_INFORMATION_EX
    PartitionStyle As Long
    UnknowData As Long ' ??? What a fucking 4 bytes is returned here in structure??? No any information :((( But other members in struct is SHIFTED by 4 bytes
    StartingOffset As Currency
    PartitionLength As Currency
    PartitionNumber As Long
    RewritePartition As Byte
    UnusedBytesForAlignTo4(0 To 2) As Byte ' ??? Another fucking ALIGNING trick from Microsoft - data must be ALIGNED to 4  :((((
    PARTITION_INFO_GPT_AND_MBR As PARTITION_INFORMATION_GPT ' Here must be a UNION that in Visual Basic is IMPOSSIBLE!!! As result we MUST use largest member from UNION - GPT info
End Type

Private Type TEST_EX
    Data(0 To 199) As Byte
End Type

'typedef struct _PARTITION_INFORMATION {
'  LARGE_INTEGER StartingOffset;
'  LARGE_INTEGER PartitionLength;
'  DWORD         HiddenSectors;
'  DWORD         PartitionNumber;
'  BYTE          PartitionType;
'  BOOLEAN       BootIndicator;
'  BOOLEAN       RecognizedPartition;
'  BOOLEAN       RewritePartition;
'} PARTITION_INFORMATION, *PPARTITION_INFORMATION;

Private Type PARTITION_INFORMATION
    StartingOffset As Currency
    PartitionLength As Currency
    HiddenSectors As Long
    PartitionNumber As Long
    PartitionType As Byte
    BootIndicator As Byte
    RecognizedPartition As Byte
    RewritePartition As Byte
End Type

'typedef struct _PartitionInfo
'{
'   PARTITION_STYLE PartitionStyle;
'   LARGE_INTEGER PartitionLength;
'   BYTE PartitionType;
'   BOOLEAN BootIndicator;
'   BOOLEAN Removable;
'   STORAGE_BUS_TYPE BusType;
'} PartitionInfo,*PPartitionInfo;

Private Type PartitionInfo
    PartitionStyle As PARTITION_STYLE
    PartitionLength As Currency
    PartitionType As Byte
    BootIndicator As Byte
    Removable As Byte
    'UnusedByteForAlignTo4 As Byte ' The same fucking trick - ALIGN data in struct to 4 bytes :(((
    BusType As STORAGE_BUS_TYPE
    ' We dont need pass this structure to any system DLL so we can add own info
    GPTInfo As PARTITION_INFORMATION_GPT
    DeviceName As String
    BusTypeString As String
End Type

Private Declare Sub RtlInitUnicodeString Lib "ntdll.dll" (DestinationString As UNICODE_STRING, ByVal SourceString As Long)

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
        ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
        lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
        ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
        (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, _
        lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
       
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hdevice As Long, ByVal dwIoControlCode As Long, _
        lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
        lpBytesReturned As Long, lpOverlapped As Any) As Byte
       
Private Declare Function NtCreateFile Lib "ntdll.dll" (FileHandle As Long, ByVal DesiredAccess As Long, _
        ObjectAttributes As OBJECT_ATTRIBUTES, IoStatusBlock As IO_STATUS_BLOCK, AllocationSize As LARGE_INTEGER, _
        ByVal FileAttributes As Long, ByVal ShareAccess As Long, ByVal CreateDisposition As Long, _
        ByVal CreateOptions As Long, EaBuffer As Any, ByVal EaLength As Long) As Long

Private Declare Function NtOpenSymbolicLinkObject Lib "ntdll.dll" _
        (LinkHandle As Long, ByVal DesiredAccess As Long, ObjAttr As OBJECT_ATTRIBUTES) As Long

Private Declare Function NtQuerySymbolicLinkObject Lib "ntdll.dll" _
        (ByVal LinkHandle As Long, LinkTarget As UNICODE_STRING, RetLength As Long) As Long
       
Private Declare Function FindFirstVolume Lib "kernel32" Alias "FindFirstVolumeA" _
            (ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long

Private Declare Function FindNextVolume Lib "kernel32" Alias "FindNextVolumeA" _
            (ByVal hFindVolume As Long, ByVal lpszVolumeName As Any, ByVal cchBufferLength As Long) As Byte

Private Declare Function FindVolumeClose Lib "kernel32" (ByVal hFindVolume As Long) As Byte
       
Private Declare Function GetVolumePathNamesForVolumeName Lib "kernel32" Alias "GetVolumePathNamesForVolumeNameA" _
        (ByVal lpszVolumeName As String, ByVal lpszVolumePathNames As String, _
        ByVal cchBufferLength As Long, lpcchReturnLength As Long) As Byte

Private Function CTL_CODE(ByVal lngDevFileSys As Long, ByVal lngFunction As Long, ByVal lngMethod As Long, ByVal lngAccess As Long) As Long
        CTL_CODE = (lngDevFileSys * (2 ^ 16)) Or (lngAccess * (2 ^ 14)) Or (lngFunction * (2 ^ 2)) Or lngMethod
End Function

'#define IOCTL_DISK_GET_PARTITION_INFO   CTL_CODE(IOCTL_DISK_BASE, 0x0001, METHOD_BUFFERED, FILE_READ_ACCESS)
Private Function IOCTL_DISK_GET_PARTITION_INFO() As Long
    IOCTL_DISK_GET_PARTITION_INFO = CTL_CODE(IOCTL_DISK_BASE, &H1, METHOD_BUFFERED, FILE_READ_ACCESS)
End Function

'#define IOCTL_DISK_GET_PARTITION_INFO_EX    CTL_CODE(IOCTL_DISK_BASE, 0x0012, METHOD_BUFFERED, FILE_ANY_ACCESS)
Private Function IOCTL_DISK_GET_PARTITION_INFO_EX() As Long
    IOCTL_DISK_GET_PARTITION_INFO_EX = CTL_CODE(IOCTL_DISK_BASE, &H12, METHOD_BUFFERED, FILE_ANY_ACCESS)
End Function

'#define IOCTL_STORAGE_QUERY_PROPERTY                CTL_CODE(IOCTL_STORAGE_BASE, 0x0500, METHOD_BUFFERED, FILE_ANY_ACCESS)
Private Function IOCTL_STORAGE_QUERY_PROPERTY() As Long
    IOCTL_STORAGE_QUERY_PROPERTY = CTL_CODE(IOCTL_STORAGE_BASE, &H500, METHOD_BUFFERED, FILE_ANY_ACCESS)
End Function

Public Function GetDrivesInfo() As Variant
    Dim i As Long, j As Long, sTmp As String
    Dim dev_num As Integer, part_num As Integer, part_info As String
    Dim sDiskInfo() As String
    Dim sDev As String, partition_info As PartitionInfo
    Dim lTotalMB As Long, lFreeMB As Long, lAvailMB As Long, lUsedMB As Long
    Dim VolumeName As String, VolumeFS As String, VolSerial As String
    Dim VolumesList() As String, VolumesListIsGood As Boolean, VolumeMatched As String
    Dim VolumeUsed() As Boolean
    Dim LettersAndFoldersForVolume() As String, lMaxInfoIndex As Long
   
    On Error Resume Next
   
    ReDim sDiskInfo(0 To MAX_DISK_INFO_PARAMS, 0 To 0)
   
   'log.WriteLog "GetDrivesInfo: retrieving volumes list..."
   
    VolumesListIsGood = GetVolumesList(VolumesList)
    If VolumesListIsGood Then
        ReDim VolumeUsed(LBound(VolumesList) To UBound(VolumesList))
        For i = LBound(VolumeUsed) To UBound(VolumeUsed)
            VolumeUsed(i) = False
        Next i
    End If
    
   'log.WriteLog "GetDrivesInfo: retrieving volumes list - " & VolumesListIsGood
   
    For dev_num = 0 To MAX_DEVICES - 1
        For part_num = 0 To MAX_PARTITIONS - 1
            sDev = "\Device\Harddisk" & CStr(dev_num) & "\Partition" & CStr(part_num)
            
            'log.WriteLog "GetDrivesInfo: cheking volume " & sDev
            
            If OpenDevice(sDev, partition_info) Then
                
                part_info = vbNullString
                'log.WriteLog "GetDrivesInfo: OpenDevice for " & sDev & " is success, analizing..."
                sTmp = GetDiskDeviceDriveLetter(sDev, part_info)
              If (sTmp <> "A") And (sTmp <> "B") And _
                                InStr(1, part_info, "CdRom", vbTextCompare) = 0 And _
                                    InStr(1, part_info, "Floppy", vbTextCompare) = 0 Then

                    'If InStr(1, part_info, "HarddiskVolume") > 0 Then
                       
                       lMaxInfoIndex = UBound(sDiskInfo, 2)
                        sDiskInfo(DISK_INFO.DriveIndex, lMaxInfoIndex) = CStr(dev_num)
                        sDiskInfo(DISK_INFO.DrivePath, lMaxInfoIndex) = sDev
                        sDiskInfo(DISK_INFO.PartitionPath, lMaxInfoIndex) = part_info
                        sDiskInfo(DISK_INFO.PartitionNumber, lMaxInfoIndex) = CStr(part_num)
                        sDiskInfo(DISK_INFO.DriveBusType, lMaxInfoIndex) = partition_info.BusTypeString
                        sDiskInfo(DISK_INFO.DriveName, lMaxInfoIndex) = partition_info.DeviceName
                        If Len(sTmp) = 1 Then sDiskInfo(DISK_INFO.PartitionLetter, lMaxInfoIndex) = sTmp
                       
                        If VolumesListIsGood Then
                            For i = LBound(VolumesList) To UBound(VolumesList)
                                VolumeMatched = vbNullString
                                If ResolveSymbolicLink(VolumesList(i), VolumeMatched) Then
                                    If VolumeMatched = part_info Then
                                        sDiskInfo(DISK_INFO.MatchedVolume, lMaxInfoIndex) = VolumesList(i)
                                        VolumeMatched = VolumesList(i)
                                        VolumeUsed(i) = True
                                        If GetVolumeLettersAndFolders(VolumeMatched, LettersAndFoldersForVolume) Then
                                            sDiskInfo(DISK_INFO.VolumeLettersAndFolders, lMaxInfoIndex) = _
                                                    Join(LettersAndFoldersForVolume, "; ")
                                            If sDiskInfo(DISK_INFO.VolumeLettersAndFolders, lMaxInfoIndex) = "; " Then _
                                                    sDiskInfo(DISK_INFO.VolumeLettersAndFolders, lMaxInfoIndex) = "None"
                                            Exit For
                                        End If
                                        'GetDiskParams sTmp, lTotalMB, lUsedMB, lFreeMB, lAvailMB, VolumeName, VolumeFS, VolSerial
                                    Else
                                        VolumeMatched = vbNullString
                                    End If
                                End If
                            Next i
                        End If
                        If sTmp = NO_DOS_LETTERS Then ' No letters for disk in "\Dosdevice\X:" style !!!
                            'sTmp = NO_DOS_LETTERS
                            sDiskInfo(DISK_INFO.PartitionLetter, lMaxInfoIndex) = sTmp
                            If Len(VolumeMatched) = 0 Then
                                sDiskInfo(DISK_INFO.TotalMegaBytes, lMaxInfoIndex) = "0"
                                sDiskInfo(DISK_INFO.UsedMegaBytes, lMaxInfoIndex) = "0"
                                sDiskInfo(DISK_INFO.FreeMegaBytes, lMaxInfoIndex) = "0"
                                sDiskInfo(DISK_INFO.AvailableMegaBytes, lMaxInfoIndex) = "0"
                                sDiskInfo(DISK_INFO.VolumeName, lMaxInfoIndex) = vbNullString
                                sDiskInfo(DISK_INFO.VolumeFileSystem, lMaxInfoIndex) = vbNullString
                                sDiskInfo(DISK_INFO.VolumeSerial, lMaxInfoIndex) = vbNullString
                            Else
                                sTmp = VolumeMatched
                            End If
                        End If
                           
                        If sTmp <> NO_DOS_LETTERS Then
                            'sDiskInfo(DISK_INFO.PartitionLetter, UBound(sDiskInfo, 2)) = sTmp
                            GetDiskParams sTmp, lTotalMB, lUsedMB, lFreeMB, lAvailMB, VolumeName, VolumeFS, VolSerial
                            sDiskInfo(DISK_INFO.TotalMegaBytes, lMaxInfoIndex) = CStr(lTotalMB)
                            sDiskInfo(DISK_INFO.UsedMegaBytes, lMaxInfoIndex) = CStr(lUsedMB)
                            sDiskInfo(DISK_INFO.FreeMegaBytes, lMaxInfoIndex) = CStr(lFreeMB)
                            sDiskInfo(DISK_INFO.AvailableMegaBytes, lMaxInfoIndex) = CStr(lAvailMB)
                            sDiskInfo(DISK_INFO.VolumeName, lMaxInfoIndex) = VolumeName
                            sDiskInfo(DISK_INFO.VolumeFileSystem, lMaxInfoIndex) = VolumeFS
                            sDiskInfo(DISK_INFO.PartitionStyle, lMaxInfoIndex) = GetPartitionStyleName(partition_info.PartitionStyle)
                            If Len(VolSerial) = 0 Or VolSerial = "0" Or partition_info.PartitionStyle = PARTITION_STYLE_GPT Then
                                sDiskInfo(DISK_INFO.VolumeSerial, lMaxInfoIndex) = modMain.GetStringFromGUID(partition_info.GPTInfo.PartitionId)
                            Else
                                sDiskInfo(DISK_INFO.VolumeSerial, lMaxInfoIndex) = VolSerial
                            End If
                            If partition_info.PartitionStyle = PARTITION_STYLE_GPT Then
                                sDiskInfo(DISK_INFO.PartitionGPT_GUID, lMaxInfoIndex) = modMain.GetStringFromGUID(partition_info.GPTInfo.PartitionId)
                                
                                sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex) = CStr(partition_info.GPTInfo.NameGPT)
                                j = InStr(1, sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex), Chr$(0), vbBinaryCompare)
                                If j > 0 Then
                                    sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex) = _
                                    Left$(sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex), j - 1)
                                End If
                                sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex) = _
                                    Replace$(sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex), "  ", " ", 1, 20, vbBinaryCompare)
                            End If
                        End If
                       
                       
                        ReDim Preserve sDiskInfo(0 To MAX_DISK_INFO_PARAMS, 0 To lMaxInfoIndex + 1)
                        lMaxInfoIndex = lMaxInfoIndex + 1
                    'End If
                End If
            End If
        Next part_num
    Next dev_num
    
    For i = LBound(VolumeUsed) To UBound(VolumeUsed)
        If Not VolumeUsed(i) Then   ' here we have some volume(s) that NOT matched with any physical partitions. It's may be raid/spanned/dynamic etc volume
            VolumeMatched = vbNullString
            If ResolveSymbolicLink(VolumesList(i), VolumeMatched) Then
                
                part_info = vbNullString
                sTmp = GetDiskDeviceDriveLetter(VolumesList(i), part_info)
                If Len(sTmp) = 0 Then sTmp = NO_DOS_LETTERS
              If (sTmp <> "A") And (sTmp <> "B") And _
                                InStr(1, part_info, "CdRom", vbTextCompare) = 0 And _
                                    InStr(1, part_info, "Floppy", vbTextCompare) = 0 Then
                    sDiskInfo(DISK_INFO.PartitionLetter, lMaxInfoIndex) = sTmp
                
                    sDiskInfo(DISK_INFO.DriveIndex, lMaxInfoIndex) = "99"
                    sDiskInfo(DISK_INFO.DrivePath, lMaxInfoIndex) = VolumesList(i)
                    sDiskInfo(DISK_INFO.PartitionPath, lMaxInfoIndex) = part_info
                    sDiskInfo(DISK_INFO.PartitionNumber, lMaxInfoIndex) = "0"
                    sDiskInfo(DISK_INFO.MatchedVolume, lMaxInfoIndex) = VolumesList(i)
                    VolumeMatched = VolumesList(i)
                    If GetVolumeLettersAndFolders(VolumeMatched, LettersAndFoldersForVolume) Then
                        sDiskInfo(DISK_INFO.VolumeLettersAndFolders, lMaxInfoIndex) = _
                                    Join(LettersAndFoldersForVolume, "; ")
                        If sDiskInfo(DISK_INFO.VolumeLettersAndFolders, lMaxInfoIndex) = "; " Then _
                                    sDiskInfo(DISK_INFO.VolumeLettersAndFolders, UBound(sDiskInfo, 2)) = "None"
                    End If
                    GetDiskParams VolumeMatched, lTotalMB, lUsedMB, lFreeMB, lAvailMB, VolumeName, VolumeFS, VolSerial
                    sDiskInfo(DISK_INFO.TotalMegaBytes, lMaxInfoIndex) = CStr(lTotalMB)
                    sDiskInfo(DISK_INFO.UsedMegaBytes, lMaxInfoIndex) = CStr(lUsedMB)
                    sDiskInfo(DISK_INFO.FreeMegaBytes, lMaxInfoIndex) = CStr(lFreeMB)
                    sDiskInfo(DISK_INFO.AvailableMegaBytes, lMaxInfoIndex) = CStr(lAvailMB)
                    sDiskInfo(DISK_INFO.VolumeName, lMaxInfoIndex) = VolumeName
                    sDiskInfo(DISK_INFO.VolumeFileSystem, lMaxInfoIndex) = VolumeFS
                    sDiskInfo(DISK_INFO.PartitionStyle, lMaxInfoIndex) = GetPartitionStyleName(partition_info.PartitionStyle)
                    If Len(VolSerial) = 0 Or VolSerial = "0" Or partition_info.PartitionStyle = PARTITION_STYLE_GPT Then
                        sDiskInfo(DISK_INFO.VolumeSerial, lMaxInfoIndex) = modMain.GetStringFromGUID(partition_info.GPTInfo.PartitionId)
                    Else
                        sDiskInfo(DISK_INFO.VolumeSerial, lMaxInfoIndex) = VolSerial
                    End If
                    If partition_info.PartitionStyle = PARTITION_STYLE_GPT Then
                        sDiskInfo(DISK_INFO.PartitionGPT_GUID, lMaxInfoIndex) = modMain.GetStringFromGUID(partition_info.GPTInfo.PartitionId)
                                
                        sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex) = CStr(partition_info.GPTInfo.NameGPT)
                        j = InStr(1, sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex), Chr$(0), vbBinaryCompare)
                        If j > 0 Then
                            sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex) = _
                            Left$(sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex), j - 1)
                        End If
                        sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex) = _
                                    Replace$(sDiskInfo(DISK_INFO.PartitionGPT_Name, lMaxInfoIndex), "  ", " ", 1, 20, vbBinaryCompare)
                    End If
                    ReDim Preserve sDiskInfo(0 To MAX_DISK_INFO_PARAMS, 0 To lMaxInfoIndex + 1)
                    lMaxInfoIndex = lMaxInfoIndex + 1
                 End If
            End If
        End If
    Next i
    
    ReDim Preserve sDiskInfo(0 To MAX_DISK_INFO_PARAMS, 0 To lMaxInfoIndex - 1)
    'ReDim Preserve sDiskInfo(0 To UBound(sDiskInfo) - 1)
    GetDrivesInfo = sDiskInfo
    Err.Clear
    On Error GoTo 0
End Function

Private Function ResolveSymbolicLink(ByRef sym_link As String, ByRef target_str As String) As Boolean

   Dim hlink As Long, i As Long
   Dim ObjectAttributes As OBJECT_ATTRIBUTES
   Dim fulldevlink As UNICODE_STRING
   Dim target As UNICODE_STRING
   Dim ret_code As Long
   Dim ReturnedLength As Long
   Dim tmpStr As String
   Dim target_ptr As String
      
    On Error Resume Next
   
   If InStr(1, sym_link, "\Volume{", vbBinaryCompare) = 0 Then
        tmpStr = sym_link & Chr$(0)
   Else
        tmpStr = sym_link
        tmpStr = Replace$(tmpStr, "\\?\", "\??\", 1, 1, vbBinaryCompare)
        tmpStr = Replace$(tmpStr, "}\", "}", 1, 1, vbBinaryCompare)
        tmpStr = tmpStr & Chr$(0)
   End If
   RtlInitUnicodeString fulldevlink, StrPtr(tmpStr)
   ObjectAttributes.Length = Len(ObjectAttributes)
   ObjectAttributes.Attributes = OBJ_KERNEL_HANDLE Or OBJ_CASE_INSENSITIVE
   ObjectAttributes.ObjectName = VarPtr(fulldevlink)
   hlink = 0
   ret_code = NtOpenSymbolicLinkObject(hlink, GENERIC_READ, ObjectAttributes)
   If ret_code <> 0 Then
      If hlink <> 0 Then CloseHandle hlink
      ResolveSymbolicLink = False
      Exit Function
    End If
   
    target_ptr = String$(MAX_PATH, " ")
    RtlInitUnicodeString target, StrPtr(target_ptr)
   target.Length = 0
   target.MaximumLength = MAX_PATH
   ret_code = NtQuerySymbolicLinkObject(hlink, target, ReturnedLength)
   If ret_code <> 0 Then
      If hlink <> 0 Then CloseHandle hlink
      Err.Clear
      On Error GoTo 0
      ResolveSymbolicLink = False
      Exit Function
   End If
   ret_code = InStr(1, target_ptr, Chr$(0))
   If ret_code > 0 Then target_ptr = Left$(target_ptr, ret_code - 1)
   target_str = target_ptr
   ResolveSymbolicLink = True
   If hlink <> 0 Then CloseHandle hlink
   
   'log.WriteLog "ResolveSymbolicLink: analizing " & sym_link & " success, result: " & target_str
      Err.Clear
      On Error GoTo 0
End Function

Private Function GetDiskDeviceDriveLetter(ByRef devName As String, ByRef partName As String) As String
    Dim i As Integer
    Dim link As String
    Dim target As String
    Dim device As String
    Dim device_letter As String
    Dim DeviceName As String
    On Error Resume Next
    DeviceName = devName
    link = String$(MAX_PATH, " ")
    target = String$(MAX_PATH, " ")
    device = String$(MAX_PATH, " ")

    If Not ResolveSymbolicLink(DeviceName, device) Then
       device = DeviceName
    End If
    partName = device
    GetDiskDeviceDriveLetter = NO_DOS_LETTERS
    ' Asc("A") = 65, Asc("Z")90
    For i = 65 To 90
      device_letter = Chr$(i)
      target = String$(MAX_PATH, " ")
      link = "\DosDevices\" & device_letter & ":"
       ResolveSymbolicLink link, target
        If StrComp(device, target) = 0 Then
            GetDiskDeviceDriveLetter = device_letter
            Exit For
        End If
    Next i
    Err.Clear
    On Error GoTo 0
End Function

Private Function OpenDevice(ByRef devName As String, ByRef ppin As PartitionInfo) As Boolean
    Dim DeviceName As String
    Dim fulldevlink As UNICODE_STRING
    Dim ObjectAttributes As OBJECT_ATTRIBUTES
    Dim hdevice As Long, ret_code As Long
    Dim IoStatus As IO_STATUS_BLOCK
    Dim LargeInt As LARGE_INTEGER
    Dim ttest As TEST_EX
    Dim pinfex As PARTITION_INFORMATION_EX
    Dim BytesReturned As Long, ppi As PartitionInfo ', pinf As PARTITION_INFORMATION
    'Dim pinf_gpt As PARTITION_INFORMATION_GPT
    Dim pinf_mbr As PARTITION_INFORMATION_MBR
    Dim spq As STORAGE_PROPERTY_QUERY
    Dim sdd As STORAGE_DEVICE_DESCRIPTOR
    
    On Error Resume Next
    'log.WriteLog "OpenDevice: trying to open " & devName & "..."
    
    DeviceName = devName & Chr$(0)
    RtlInitUnicodeString fulldevlink, StrPtr(DeviceName)
   
    ObjectAttributes.Length = LenB(ObjectAttributes)
    ObjectAttributes.RootDirectory = 0&
    ObjectAttributes.SecurityDescriptor = 0&
    ObjectAttributes.SecurityQualityOfService = 0&
    ObjectAttributes.Attributes = OBJ_KERNEL_HANDLE Or OBJ_CASE_INSENSITIVE
    ObjectAttributes.ObjectName = VarPtr(fulldevlink)
    hdevice = 0
    ret_code = NtCreateFile(hdevice, SYNCHRONIZE Or FILE_READ_ATTRIBUTES, ObjectAttributes, IoStatus, LargeInt, _
        0, FILE_SHARE_READ Or FILE_SHARE_WRITE, FILE_OPEN, FILE_SYNCHRONOUS_IO_NONALERT, 0, 0)
    OpenDevice = False
    If ret_code = 0 Then OpenDevice = True
    
    'log.WriteLog "OpenDevice: NtCreateFile ret_code=" & ret_code & " (0 - success, other fail)"
    
    If hdevice = 0 Then OpenDevice = False: Exit Function
    
    If DeviceIoControl(hdevice, IOCTL_DISK_GET_PARTITION_INFO_EX, 0, 0, pinfex, LenB(pinfex), BytesReturned, 0) <> 0 Then
        
        'log.WriteLog "OpenDevice: DeviceIoControl with IOCTL_DISK_GET_PARTITION_INFO_EX - success"
        
       ppi.PartitionStyle = pinfex.PartitionStyle
       ppi.PartitionLength = pinfex.PartitionLength
       If pinfex.PartitionStyle = PARTITION_STYLE_MBR Then
          LSet pinf_mbr = pinfex.PARTITION_INFO_GPT_AND_MBR
          ppi.BootIndicator = pinf_mbr.BootIndicator
          ppi.PartitionType = pinf_mbr.PartitionType
       Else
          '  BytesReturned = 0
            ppi.GPTInfo = pinfex.PARTITION_INFO_GPT_AND_MBR
          'If DeviceIoControl(hdevice, IOCTL_DISK_GET_PARTITION_INFO, 0, 0, pinf, LenB(pinf), BytesReturned, 0) <> 0 Then
          '  LSet pinf = ttest
          '  ppi.BootIndicator = pinf.BootIndicator
          '  ppi.PartitionType = pinf.PartitionType
          'End If
       End If
    End If
    
    spq.PropertyId = StorageDeviceProperty
    spq.QueryType = PropertyStandardQuery
    BytesReturned = 0
    If DeviceIoControl(hdevice, IOCTL_STORAGE_QUERY_PROPERTY, spq, LenB(spq), sdd, LenB(sdd), BytesReturned, ByVal 0&) <> 0 Then

        ppi.Removable = sdd.RemovableMedia
        ppi.BusType = sdd.BusType
        ppi.BusTypeString = GetStorageBusName(ppi.BusType)
        'ppi.DeviceName = GetStorageParamFromRAW(sdd.RawDeviceProperties)
        
    End If
    ppin = ppi
    
    If hdevice <> 0 Then CloseHandle hdevice
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetStorageParamFromRAW(ByRef RAWParamArray, Optional ByVal StartOffset) As String
    Dim sRes As String, bRes() As Byte, i As Long, l1 As Long, l2 As Long
    If Not IsArray(RAWParamArray) Then Exit Function
    bRes = RAWParamArray
    l1 = LBound(bRes)
    If Not IsMissing(StartOffset) Then l1 = StartOffset
    l2 = UBound(bRes)
    sRes = vbNullString
    i = l1
    Do While bRes(i) = 0
        i = i + 1
    Loop
    l1 = i
    For i = l1 To l2
        If bRes(i) = 0 Then Exit For
        sRes = sRes & Chr$(bRes(i))
    Next i
    i = InStr(sRes, "   ")
    If i > 0 Then sRes = Left$(sRes, i - 1)
    GetStorageParamFromRAW = sRes
End Function

Private Function GetDiskParams(ByRef DiskLetter As String, ByRef TotalMB As Long, _
                ByRef UsedMB As Long, ByRef FreeMB As Long, ByRef AvailMB As Long, _
                ByRef VolName As String, ByRef FileSysName As String, ByRef VolumeSerialNumber As String)
    Dim BytesFreeToCalller As Currency, TotalBytes As Currency
    Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
    Dim RootPathName As String
    Dim Volume_Name As String
    Dim File_System_Name As String
    Dim Info_Status As Long ', MaxComponentLength As Long
    'Dim File_System_Flags As Long, Serial_Number As Long
    Dim Serial_Number As Long
    Dim pos As Long
   
    On Error Resume Next
    If Len(DiskLetter) = 1 Then
        RootPathName = DiskLetter & ":\"
    Else
        RootPathName = DiskLetter
    End If
    
    'log.WriteLog "GetDiskParams: Starting GetDiskFreeSpaceEx for drive " & RootPathName & "..."
    Call GetDiskFreeSpaceEx(RootPathName, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    'log.WriteLog "GetDiskParams: GetDiskFreeSpaceEx for drive " & RootPathName & " - SUCCESS"
    BytesFreeToCalller = BytesFreeToCalller * 10000
    TotalBytes = TotalBytes * 10000
    TotalFreeBytes = TotalFreeBytes * 10000
   
    TotalMB = CLng(TotalBytes / 1024 / 1024)
    FreeMB = CLng(TotalFreeBytes / 1024 / 1024)
    AvailMB = CLng(BytesFreeToCalller / 1024 / 1024)
    UsedMB = TotalMB - FreeMB
    'show the results, multiplying the returned
    'value by 10000 to adjust for the 4 decimal
    'places that the currency data type returns.
    'Me.Print " Total Number Of Bytes:", Format$(TotalBytes * 10000, "###,###,###,##0") & " bytes"
    'Me.Print " Total Free Bytes:", Format$(TotalFreeBytes * 10000, "###,###,###,##0") & " bytes"
    'Me.Print " Free Bytes Available:", Format$(BytesFreeToCalller * 10000, "###,###,###,##0") & " bytes"
    'Me.Print " Total Space Used :", Format$((TotalBytes - TotalFreeBytes) * 10000, "###,###,###,##0") & " bytes"
   
    Volume_Name = String$(255, Chr$(0))
    File_System_Name = String$(255, Chr$(0))

    ' Get the volume information.
    'log.WriteLog "GetDiskParams: Starting GetVolumeInformation for drive " & RootPathName & "..."
    Info_Status = GetVolumeInformation(RootPathName, Volume_Name, 255, Serial_Number, _
        0, 0, File_System_Name, 255)
    If Info_Status <> 0 Then
        'log.WriteLog "GetDiskParams: GetVolumeInformation for drive " & RootPathName & " - SUCCESS"
        pos = InStr(Volume_Name, vbNullChar)
        If pos > 0 Then VolName = Left$(Volume_Name, pos - 1) Else VolName = Volume_Name
        pos = InStr(File_System_Name, vbNullChar)
        If pos > 0 Then FileSysName = Left$(File_System_Name, pos - 1) Else FileSysName = File_System_Name
        VolumeSerialNumber = Hex$(Serial_Number)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetVolumesList(ByRef resultArray As Variant) As Boolean
    Dim volHandle As Long, volPath As String, volLength As Long, lPos As Long
    Dim resArray() As String
    On Error Resume Next
    volPath = String$(255, Chr$(0))
    volLength = 255
    'log.WriteLog "GetVolumesList: FindFirstVolume started"
    volHandle = FindFirstVolume(volPath, volLength)
    If volHandle = 0 Then
        GetVolumesList = False
        Exit Function
    End If
    lPos = InStr(volPath, Chr$(0))
    If lPos > 0 Then volPath = Left$(volPath, lPos - 1)
        
        'log.WriteLog "GetVolumesList: FindFirstVolume returned " & volPath

    ReDim resArray(0 To 0)
    resArray(0) = volPath
    volPath = String$(255, Chr$(0))
    volLength = 255
    
'log.WriteLog "GetVolumesList: FindNextVolume started"

    Do Until FindNextVolume(volHandle, volPath, volLength) = 0
        lPos = InStr(volPath, Chr$(0))
        If lPos > 0 Then volPath = Left$(volPath, lPos - 1)
        'log.WriteLog "GetVolumesList: FindNextVolume returned " & volPath
        ReDim Preserve resArray(0 To UBound(resArray) + 1)
        resArray(UBound(resArray)) = volPath
    Loop
    FindVolumeClose volHandle
    resultArray = resArray
    GetVolumesList = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetVolumeLettersAndFolders(ByRef sVolume As String, ByRef resultArray As Variant) As Boolean
    Dim pathsList As String, pathsLength As Long, pathVolume As String, lRetLentgh As Long
    Dim resArray() As String
    On Error Resume Next
    pathVolume = sVolume
    pathsList = String$(255, Chr$(0))
    pathsLength = 255
    GetVolumeLettersAndFolders = False
    If GetVolumePathNamesForVolumeName(pathVolume, pathsList, pathsLength, lRetLentgh) <> 0 Then
        pathsList = Left$(pathsList, lRetLentgh)
        pathsList = Replace$(pathsList, Chr$(0) & Chr$(0), vbNullString, 1, lRetLentgh, vbBinaryCompare)
        resArray = Split(pathsList, Chr$(0))
        resultArray = resArray
        GetVolumeLettersAndFolders = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetStorageBusName(ByVal BusType As Long) As String
    Dim sRes As String
    Select Case BusType
        Case BusTypeScsi
            sRes = "SCSI"
        Case BusTypeAtapi
            sRes = "ATAPI"
        Case BusTypeAta
            sRes = "ATA"
        Case BusType1394
            sRes = "Firewire"
        Case BusTypeSsa
            sRes = "SSA"
        Case BusTypeFibre
            sRes = "FibreChannel"
        Case BusTypeUsb
            sRes = "USB"
        Case BusTypeRAID
            sRes = "RAID"
        Case BusTypeiSCSI
            sRes = "iSCSI"
        Case BusTypeSas
            sRes = "SAS"
        Case BusTypeSata
            sRes = "SATA"
        Case Else
            sRes = "Unknown"
    End Select
    GetStorageBusName = sRes
End Function

Private Function GetPartitionStyleName(ByVal PartitionType As Long) As String
    Dim sRes As String
    Select Case PartitionType
        Case PARTITION_STYLE_MBR
            sRes = "MBR"
        Case PARTITION_STYLE_GPT
            sRes = "GPT"
        Case PARTITION_STYLE_RAW
            sRes = "RAW"
        Case Else
            sRes = "Unknown"
    End Select
    GetPartitionStyleName = sRes
End Function



