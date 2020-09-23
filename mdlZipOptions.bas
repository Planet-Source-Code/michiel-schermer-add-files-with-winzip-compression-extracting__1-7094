Attribute VB_Name = "mdlZipOptions"
' This program is made by a M. Schermer from the Netherlands
' I'm a prof. programmer for a lot of compagny's in the Netherlands and was in America
' a short time ago to program for a compagny that was making radars and that kind of stuff
' So you know from where you got this code

' Look in a couple of days on www.Planet-Source-Code.com for an update of this version.
' Sorry if there are some spelling faults in my notes or that there was left over some Dutch
' documentation

' Good luck with the code - Michiel.Schermer@Bit-ic.nl
    
    Public Options As ZipOptions
        Public Type ZipOptions
            ActionToDo As ActionZIP
            Options As OptionsZIP
            Compression As CompressionZIP
            PassWord As String
            FilesToAdd As FilesToAddZIP
            IfFileAlreadyExists As IfFileExists
        End Type
    
    
    Public Enum ActionZIP
        Add = 1                        'Add to archive
        Freshen = 2                    'Freshen archive
        Update = 3                     'Update archive
        Move = 4                       'Move to archive
    End Enum

    Public Enum OptionsZIP
        Recurse_Directories = 1        'Recurse Directories
        Save_Extra_Directory_Info = 2  'Save Extra Directory Info
    End Enum

    Public Enum CompressionZIP
        eXtra = 1                      'Extra
        Normal = 2                     'Normal
        Fast = 3                       'Fast
        Super_Fast = 4                 'Super fast
        No_Compression = 5             'No compression
    End Enum

    Public Enum PassWordZIP
        PassWord = 1                   'Password protection.
    End Enum

    Public Enum FilesToAddZIP
        AddHiddenSystem = 1            'Add also hidden and system files to archive
        DoNotAddHiddenSystem = 1
    End Enum

    Public Enum IfFileExists
        AlwaysOverwrite = 1
        NotOverwrite = 2
    End Enum
