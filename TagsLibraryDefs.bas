Attribute VB_Name = "modTagsLibrary"

' Tags Library Visual Basic module
' Copyright (c) 2014 3delite.
'
' See the Tags Library ReadMe.txt file for more detailed documentation

' NOTE: Use "StrPtr(string)" for functions expecting a wide string pointer.

' NOTE: Use the LPWSTRtoBSTR function to convert "char *" to VB "String".

Global Const TagsLibraryName = "TagsLib.dll"

Global Const TAGSLIBRARY_SUCCESS                                             = 0;
Global Const TAGSLIBRARY_ERROR                                               = $FFFF;
Global Const TAGSLIBRARY_ERROR_NO_TAG_FOUND                                  = 1;
Global Const TAGSLIBRARY_ERROR_FILENOTFOUND                                  = 2;
Global Const TAGSLIBRARY_ERROR_EMPTY_TAG                                     = 3;
Global Const TAGSLIBRARY_ERROR_EMPTY_FRAMES                                  = 4;
Global Const TAGSLIBRARY_ERROR_OPENING_FILE                                  = 5;
Global Const TAGSLIBRARY_ERROR_READING_FILE                                  = 6;
Global Const TAGSLIBRARY_ERROR_WRITING_FILE                                  = 7;
Global Const TAGSLIBRARY_ERROR_CORRUPT                                       = 8;
Global Const TAGSLIBRARY_ERROR_NOT_SUPPORTED_VERSION                         = 9;
Global Const TAGSLIBRARY_ERROR_NOT_SUPPORTED_FORMAT                          = 10;
Global Const TAGSLIBRARY_ERROR_BASS_NOT_LOADED                               = 11;
Global Const TAGSLIBRARY_ERROR_BASS_ChannelGetTags_NOT_FOUND                 = 12;
Global Const TAGSLIBRARY_ERROR_DOESNT_FIT                                    = 13;
Global Const TAGSLIBRARY_ERROR_NEED_EXCLUSIVE_ACCESS                         = 14;
Global Const TAGSLIBRARY_ERROR_WMATAGLIBRARY_COULDNTLOADDLL                  = 15;
Global Const TAGSLIBRARY_ERROR_WMATAGLIBRARY_COULDNOTCREATEMETADATAEDITOR    = 16;
Global Const TAGSLIBRARY_ERROR_WMATAGLIBRARY_COULDNOTQIFORIWMHEADERINFO3     = 17;
Global Const TAGSLIBRARY_ERROR_WMATAGLIBRARY_COULDNOTQUERY_ATTRIBUTE_COUNT   = 18;
Global Const TAGSLIBRARY_ERROR_MP4TAGLIBRARY_UPDATE_stco                     = 19;
Global Const TAGSLIBRARY_ERROR_MP4TAGLIBRARY_UPDATE_co64                     = 20;

Global Const TAGSLIBRARY_PADDING_SIZE_TO_WRITE                               = 1;
Global Const TAGSLIBRARY_PARSE_OGG_PLAYTIME                                  = 2;
Global Const TAGSLIBRARY_PARSE_ID3v2_AUDIO_ATTRIBUTES                        = 3;

type HTags as Long

Enum TTagType
    ttNone
    ttAutomatic
    ttAPEv2
    ttFlac
    ttID3v1
    ttID3v2
    ttMP4
    ttOpusVorbis
    ttWAV
    ttWMA
End Enum

Enum TTagPictureFormat
    tpfUnknown
    tpfJPEG
    tpfPNG
    tpfBMP
    tpfGIF
End Enum

type TTagPriority
    Values As TTagType * 9
End Type

Enum TExtTagType
    ettUnknown
    ettTXXX
    ettWXXX
End Enum

type TExtTag
    Name As Long
    Value As Long
    ValueSize As Integer
    Language As Long
    Description As Long
    ExtTagType As TExtTagType
    Index As Integer
End Type

type TCoverArtData
    Name As Long
    Data As Long
    DataSize As Int64
    Description As Long
    CoverType As Long
    MIMEType As Long
    PictureFormat As TTagPictureFormat
    Width As Long
    Height As Long
    ColorDepth As Long
    NoOfColors As Long
    ID3v2TextEncoding As Integer
    Index As Integer
End Type

type TTagData
    Name As Long
    Data As Long
    DataSize As Int64
    DataType As Integer
End Type

type TCARTPostTimer
    Usage As Long
    Value As Long
End Type

Enum TAudioType
    atAutomatic
    atFlac
    atMPEG
    atDSDDSF
    atWAV
    atAIFF
    atMP4
    atOpus
    atVorbis
    atWMA
End Enum

Enum TMPEGVersion
    tmpegvUnknown
    tmpegv1
    tmpegv2
    tmpegv25
End Enum

Enum TMPEGLayer
    tmpeglUnknown
    tmpegl1
    tmpegl2
    tmpegl3
End Enum

Enum TMPEGChannelMode
    tmpegcmUnknown
    tmpegcmMono
    tmpegcmDualChannel
    tmpegcmJointStereo
    tmpegcmStereo
End Enum

Enum TMPEGModeExtension
    tmpegmeUnknown
    tmpegmeNone
    tmpegmeIntensity
    tmpegmeMS
    tmpegmeIntensityMS
End Enum

Enum TMPEGEmphasis
    tmpegeUnknown
    tmpegeNo
    tmpege5015
    tmpegeCCITJ17
End Enum

type TMPEGAudioAttributes
    Position As Int64                	'* Position of header in bytes
    Header As Long                  	'* The Headers bytes
    FrameSize As Integer             	'* Frame's length
    Version As TMPEGVersion          	'* MPEG Version
    Layer As TMPEGLayer              	'* MPEG Layer
    CRC As Long                  	'* Frame has CRC
    BitRate As Long                 	'* Frame's bitrate
    SampleRate As Long            	'* Frame's sample rate
    Padding As Long              	'* Frame is padded
    _Private As Long             	'* Frame's private bit is set
    ChannelMode As TMPEGChannelMode  	'* Frame's channel mode
    ModeExtension As TMPEGModeExtension '* Joint stereo only
    Copyrighted As Long;          	'* Frame's Copyright bit is set
    Original As Long;             	'* Frame's Original bit is set
    Emphasis As TMPEGEmphasis        	'* Frame's emphasis mode
    VBR As Long                  	'* Stream is probably VBR
    FrameCount As Int64              	'* Total number of MPEG frames (by header)
    Quality As Integer               	'* MPEG quality
    Bytes As Int64                   	'* Total bytes
End Type

type TFlacAudioAttributes
    Channels As Long
    SampleRate As Integer
    BitsPerSample As Long
    SampleCount As Int64
    Playtime As Double       		' Duration (seconds)
    Ratio As Double          		' Compression ratio (%)
    ChannelMode As Long
    Bitrate As Integer
End Type

Enum TDSFChannelType
    dsfctUnknown
    dsfctMono
    dsfctStereo
    dsfct3Channels
    dsfctQuad
    dsfct4Channels
    dsfct5Channels
    dsfct51Channels
End Enum

type TDSFAudioAttributes
    FormatVersion As Long
    FormatID As Long
    ChannelType As TDSFChannelType
    ChannelNumber As Long
    SamplingFrequency As Long
    BitsPerSample As Long
    SampleCount As Int64
    BlockSizePerChannel As Long
    PlayTime As Double
    Bitrate As Integer
End Type

type TOpusAudioAttributes
    BitstreamVersion As Long     	'{ Bitstream version number }
    ChannelCount As Long           	'      { Number of channels }
    PreSkip As Long
    SampleRate As Long               	'        { Sample rate (hz) }
    OutputGain As Long
    MappingFamily As Long          	'                 { 0,1,255 }
    PlayTime As Double
    SampleCount As Int64
    Bitrate As Integer
End Type

type TVorbisAudioAttributes
    BitstreamVersion As Byte * 4   	'{ Bitstream version number }
    ChannelMode As Long            	'      { Number of channels }
    SampleRate As Integer               '        { Sample rate (hz) }
    BitRateMaximal As Integer           '    { Bit rate upper limit }
    BitRateNominal As Integer           '        { Nominal bit rate }
    BitRateMinimal As Integer           '    { Bit rate lower limit }
    BlockSize As Long              	'{ Coded size for small and long blocks }
    PlayTime As Double
    SampleCount As Int64
    Bitrate As Integer
End Type

type TWAVEAudioAttributes
    FormatTag As Long                   ' format type
    Channels As Long                    ' number of channels (i.e. mono, stereo, etc.)
    SamplesPerSec As Long              	' sample rate
    AvgBytesPerSec As Long             	' for buffer estimation
    BlockAlign As Long                  ' block size of data
    BitsPerSample As Long               ' number of bits per sample of mono data
    PlayTime As Double
    SampleCount As Int64
    cbSize As Long	                ' Size of the extension: 22
    ValidBitsPerSample As Long	        ' at most 8 *  M
    ChannelMask As Long	               	' Speaker position mask: 0
    SubFormat As Byte * 16
    Bitrate As Integer
End Type

type TAIFFAttributes
    Channels As Long
    SampleCount As Long
    SampleSize As Long
    SampleRate As Double
    CompressionID As Byte * 4  		' http:'en.wikipedia.org/wiki/Audio_Interchange_File_Format
    Compression As Long
    PlayTime As Double
    BitRate As Integer
End Type

type TWMAAttributes
    PlayTime  As Double
    BitRate As Integer
End Type

type TAudioAttributes
    Channels As Long                    ' number of channels (i.e. mono, stereo, etc.)
    SamplesPerSec As Long               ' sample rate
    BitsPerSample As Long               ' number of bits per sample of mono data
    PlayTime As Double                  ' duration in seconds
    SampleCount As Int64              	' number of total samples
    Bitrate As Integer
End Type

Declare Function TagsLibrary_Create Lib "TagsLib.dll" () As HTags
Declare Function TagsLibrary_Free Lib "TagsLib.dll" (ByVal Tags As HTags) As Long
Declare Function TagsLibrary_Load Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal FileName As Long, ByVal TagType As TTagType, ByVal ParseTags As Long) As Integer
Declare Function TagsLibrary_LoadFromBASS Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Channel AsLong) As Integer
Declare Function TagsLibrary_Save Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal FileName AsLong, ByVal TagType as TTagType) As Integer
Declare Function TagsLibrary_SaveEx Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal FileName AsLong, ByVal TagType As TTagType) As Integer
Declare Function TagsLibrary_RemoveTag Lib "TagsLib.dll" (FileName As Long, ByVal TagType As TTagType) As Integer
Declare Function TagsLibrary_GetTag Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Name AsLong, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_Loaded Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_GetTagEx Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Name As Long, ByVal TagType As TTagType, ByRef ExtTag As TExtTag) As Long
Declare Function TagsLibrary_GetTagByIndexEx Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Index As Integer, ByVal TagType As TTagType, ByRef ExtTag As TExtTag) As Long
Declare Function TagsLibrary_SetTag Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Name As Long, ByVal Value As Long, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_SetTagEx Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType, ByVal ExtTag As TExtTag) As Long
Declare Function TagsLibrary_AddTag Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Name As Long, ByVal Value As Long, ByVal TagType As TTagType) As Integer
Declare Function TagsLibrary_AddTagEx Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType, ByVal ExtTag As TExtTag) As Integer
Declare Function TagsLibrary_TagCount Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType) As Integer
Declare Function TagsLibrary_DeleteTag Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Name As Long, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_DeleteTagByIndex Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Index As Integer, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_CoverArtCount Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType) As Integer
Declare Function TagsLibrary_GetCoverArt Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType, ByVal Index As Integer, ByRef CoverArt As TCoverArtData) As Long
Declare Function TagsLibrary_DeleteCoverArt Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType, ByVal Index As Integer) As Long
Declare Function TagsLibrary_SetCoverArt Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType, ByVal Index As Integer, ByRef CoverArt As TCoverArtData) As Long
Declare Function TagsLibrary_AddCoverArt Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType, ByVal CoverArt As TCoverArtData) As Integer
Declare Function TagsLibrary_SetTagLoadPriority Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagPriorities As TTagPriority) As Long
Declare Function TagsLibrary_GetTagData Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Index As Integer, ByVal TagType As TTagType, ByRef TagData As TTagData) As Long
Declare Function TagsLibrary_SetTagData Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Index As Integer, ByVal TagType As TTagType, ByVal TagData As TTagData) As Long
Declare Function TagsLibrary_GetCARTPostTimer Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Index As Integer, ByRef PostTimer As TCARTPostTimer) As Long
Declare Function TagsLibrary_SetCARTPostTimer Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Index As Integer, ByVal PostTimer As TCARTPostTimer) As Long
Declare Function TagsLibrary_ClearCARTPostTimer Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Index As Integer) As Long
Declare Function TagsLibrary_GetConfig Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Option As Integer, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_SetConfig Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Value As Long, ByVal Option As Integer, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_GetVendor Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_SetVendor Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal Vendor As Long, ByVal TagType As TTagType) As Long
Declare Function TagsLibrary_GetAudioAttributes Lib "TagsLib.dll" (ByVal Tags As HTags, ByVal AudioType As TAudioType, ByRef Attributes As Long) As Long

'This is used to find the length of a string in a pointer
Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'This is used to copy a string from a pointer
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'This is a handy utility (from the MSDN library) which converts a pointer (to a wide string) to a string
Public Function LPWSTRtoBSTR(ByVal lpwsz As Long) As String
    ' Input: a valid LPWSTR pointer lpwsz
    ' Return: a sBSTR with the same character array
    Dim cChars As Long
    ' Get number of characters in lpwsz
    cChars = lstrlenW(lpwsz)
    ' Initialize string
    LPWSTRtoBSTR = String$(cChars, 0)
    ' Copy string
    Call CopyMemory(ByVal StrPtr(LPWSTRtoBSTR), ByVal lpwsz, cChars * 2)
End Function

End Sub

