Module SharedGlobals

    '--- a messenger object
    Public MyMessenger As Messenger

    '--- a crumbler object
    Public MyCrumbler As Crumbler

    '--- a messenger object
    Public MyChatter As Chatter

    '--- command-line switches
    Public MySwitches As Switches

    '--- status message handler
    Public MyStatusMessages As StatusMessage

    'If can't write errors etc to queue, these are last ditch locations
    Public Const sNETWORK_ERROR_LOGFILE As String = "W:\Broil\Apps_dev\Marvin\ERRORS\"
    Public Const sLOCAL_ERROR_LOGFILE As String = "c:\temp\"

    'If it's not passed as a cmd line switch, use this config file
    Public Const sDEFAULT_CONFIG_FILE As String = "W:\Broil\Apps_dev\Marvin\AppConfig\AppConfig.xml"

End Module
